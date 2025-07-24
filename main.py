import logging
import argparse
import json
import os
import shutil
import atexit
from datetime import datetime
from ai_service import generate_presentation_plan
from ppt_builder.presentation import PresentationBuilder
from config import OUTPUT_DIR

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', force=True)

# 定义临时目录和输出目录
TEMP_DIR = "temp"


def cleanup_temp_dir():
    """清理临时目录。"""
    if os.path.exists(TEMP_DIR):
        try:
            shutil.rmtree(TEMP_DIR)
            logging.info(f"临时目录 '{TEMP_DIR}' 已被成功清理。")
        except OSError as e:
            logging.error(f"清理临时目录 '{TEMP_DIR}' 时出错: {e}")


def ensure_dirs_exist():
    """确保输出目录和临时目录存在。"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    os.makedirs(TEMP_DIR, exist_ok=True)


def generate_single_ppt(theme: str, num_pages: int, output_file: str, aspect_ratio: str):
    """
    为单个主题生成演示文稿。
    [已更新] 新增 aspect_ratio 参数。
    """
    logging.info(f"主题: '{theme}', 页数: {num_pages}, 宽高比: {aspect_ratio}, 输出文件: '{output_file}'")

    logging.info(f"正在请求AI为主题 '{theme}' 生成 ({aspect_ratio}) 方案...")
    # 将宽高比传递给AI服务
    plan = generate_presentation_plan(theme, num_pages, aspect_ratio)

    if plan:
        logging.info("AI方案生成成功，开始构建演示文稿。")
        try:
            full_output_path = os.path.join(OUTPUT_DIR, output_file)
            # 将宽高比传递给构建器
            builder = PresentationBuilder(plan, aspect_ratio)
            builder.build_presentation(full_output_path)
            logging.info(f"演示文稿生成完成，已保存至 {full_output_path}")
            return True
        except Exception as e:
            logging.error(f"为主题 '{theme}' 构建演示文稿失败: {e}", exc_info=True)
            return False
    else:
        logging.error(f"未能为主题 '{theme}' 生成演示文稿方案，任务中止。")
        return False


def get_formatted_filename(base_name: str) -> str:
    """根据基础名称生成带有时间戳的文件名。"""
    sanitized_base_name = base_name.removesuffix('.pptx').replace(' ', '_').replace('/', '_').replace('\\', '_')
    timestamp = datetime.now().strftime("%Y%m%d-%H%M")
    return f"{sanitized_base_name}_{timestamp}.pptx"


def main():
    """主函数，支持通过命令行参数进行单次生成，或通过配置文件进行批量生成。"""
    atexit.register(cleanup_temp_dir)
    ensure_dirs_exist()
    cleanup_temp_dir()
    ensure_dirs_exist()

    parser = argparse.ArgumentParser(description="AI PPT Generator")
    parser.add_argument("--theme", type=str, help="演示文稿的主题 (单次模式)。")
    parser.add_argument("--pages", type=int, default=10, help="演示文稿的页数。")
    parser.add_argument("--output", type=str, help="演示文稿的输出文件名 (单次模式, 无需后缀和时间戳)。")
    parser.add_argument("--batch", type=str, help="用于批量处理的JSON文件路径。")
    # --- [核心修改] 新增命令行参数 ---
    parser.add_argument(
        "--aspect_ratio",
        type=str,
        default="16:9",
        choices=["16:9", "4:3"],
        help="演示文稿的宽高比 (可选 '16:9' 或 '4:3')。"
    )
    args = parser.parse_args()

    if args.batch:
        logging.info("--- 开始批量生成PPT任务 ---")
        try:
            with open(args.batch, 'r', encoding='utf-8') as f:
                batch_tasks = json.load(f)

            total_tasks = len(batch_tasks)
            for i, task in enumerate(batch_tasks):
                logging.info(f"\n--- 正在生成第 {i + 1}/{total_tasks} 个演示文稿 ---")
                theme = task.get("theme")
                if not theme:
                    logging.warning(f"跳过任务 {i + 1}，原因：缺少'theme'。")
                    continue

                pages = task.get("pages", args.pages)
                # 优先使用任务中定义的宽高比，否则使用命令行的默认值
                aspect_ratio = task.get("aspect_ratio", args.aspect_ratio)
                base_name = task.get("output", theme)
                output_filename = get_formatted_filename(base_name)

                generate_single_ppt(theme, pages, output_filename, aspect_ratio)

        except FileNotFoundError:
            logging.error(f"批量处理文件未找到: {args.batch}")
        except json.JSONDecodeError:
            logging.error(f"解析批量处理文件JSON时出错: {args.batch}")
        except Exception as e:
            logging.error(f"批量处理过程中发生错误: {e}", exc_info=True)

    elif args.theme:
        logging.info("--- 开始单次生成PPT任务 ---")
        base_name = args.output if args.output else args.theme
        output_filename = get_formatted_filename(base_name)
        # 在单次任务中传递宽高比
        generate_single_ppt(args.theme, args.pages, output_filename, args.aspect_ratio)
    else:
        logging.warning("未指定操作。请使用 --theme 进行单次生成，或使用 --batch 进行批量处理。")
        parser.print_help()


if __name__ == "__main__":
    main()
