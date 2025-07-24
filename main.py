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


def generate_single_ppt(theme: str, num_pages: int, aspect_ratio: str):
    """
    [已更新] 为单个主题生成演示文稿，并根据内容自动命名文件。
    """
    logging.info(f"任务开始 - 主题: '{theme}', 页数: {num_pages}, 宽高比: {aspect_ratio}")

    logging.info(f"正在请求AI为主题 '{theme}' 生成 ({aspect_ratio}) 方案...")
    plan = generate_presentation_plan(theme, num_pages, aspect_ratio)

    if plan:
        logging.info("AI方案生成成功，开始构建演示文稿。")
        try:
            # --- [核心修改] 自动生成文件名 ---
            # 1. 从方案中获取设计风格
            style = plan.get('design_concept', '未知风格')
            # 2. 获取当前日期
            date_str = datetime.now().strftime("%Y%m%d")

            # 3. 清理文件名中的非法字符
            sanitized_theme = theme.replace(' ', '_').replace('/', '_').replace('\\', '_').replace(':', '：').replace(
                '《', '').replace('》', '')
            sanitized_style = style.replace(' ', '_').replace('/', '_').replace('\\', '_').replace(':', '：')

            # 4. 清理并格式化宽高比
            ratio_str = aspect_ratio.replace(':', 'x')

            # 5. 组合成最终文件名
            output_filename = f"{sanitized_theme}_{sanitized_style}_{date_str}_{ratio_str}.pptx"
            full_output_path = os.path.join(OUTPUT_DIR, output_filename)
            logging.info(f"自动生成文件名: {output_filename}")
            # --- 修改结束 ---

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


def main():
    """主函数，支持通过命令行参数进行单次生成，或通过配置文件进行批量生成。"""
    atexit.register(cleanup_temp_dir)
    ensure_dirs_exist()
    cleanup_temp_dir()
    ensure_dirs_exist()

    parser = argparse.ArgumentParser(description="AI PPT Generator")
    parser.add_argument("--theme", type=str, help="演示文稿的主题 (单次模式)。")
    parser.add_argument("--pages", type=int, default=10, help="演示文稿的页数。")
    parser.add_argument("--batch", type=str, help="用于批量处理的JSON文件路径。")
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
                aspect_ratio = task.get("aspect_ratio", args.aspect_ratio)

                # [已简化] 直接调用生成函数，无需再处理文件名
                generate_single_ppt(theme, pages, aspect_ratio)

        except FileNotFoundError:
            logging.error(f"批量处理文件未找到: {args.batch}")
        except json.JSONDecodeError:
            logging.error(f"解析批量处理文件JSON时出错: {args.batch}")
        except Exception as e:
            logging.error(f"批量处理过程中发生错误: {e}", exc_info=True)

    elif args.theme:
        logging.info("--- 开始单次生成PPT任务 ---")
        # [已简化] 直接调用生成函数
        generate_single_ppt(args.theme, args.pages, args.aspect_ratio)
    else:
        logging.warning("未指定操作。请使用 --theme 进行单次生成，或使用 --batch 进行批量处理。")
        parser.print_help()


if __name__ == "__main__":
    main()
