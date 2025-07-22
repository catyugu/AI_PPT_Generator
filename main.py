import logging
import argparse
import json
import os
from ai_service import generate_presentation_plan
from ppt_builder.presentation import PresentationBuilder
from config import OUTPUT_DIR

# 配置日志，force=True确保在任何情况下都能覆盖默认配置
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', force=True)

def ensure_output_dir_exists():
    """确保输出目录存在。"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        logging.info(f"输出目录 '{OUTPUT_DIR}' 已创建。")

def generate_single_ppt(theme: str, num_pages: int, output_file: str):
    """为单个主题生成演示文稿。"""
    logging.info(f"主题: '{theme}', 页数: {num_pages}, 输出文件: '{output_file}'")

    # 1. 从AI获取演示文稿计划
    logging.info(f"正在请求AI为主题 '{theme}' 生成方案...")
    plan = generate_presentation_plan(theme, num_pages)

    if plan:
        logging.info("AI方案生成成功，开始构建演示文稿。")
        # 2. 根据计划构建演示文稿
        try:
            # 确保输出文件路径包含输出目录
            full_output_path = os.path.join(OUTPUT_DIR, output_file)
            builder = PresentationBuilder(plan)
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
    parser = argparse.ArgumentParser(description="AI PPT Generator")
    parser.add_argument("--theme", type=str, help="演示文稿的主题 (单次模式)。")
    parser.add_argument("--pages", type=int, default=10, help="演示文稿的页数。")
    parser.add_argument("--output", type=str, help="演示文稿的输出文件名 (单次模式)。")
    parser.add_argument("--batch", type=str, help="用于批量处理的JSON文件路径。")
    args = parser.parse_args()

    ensure_output_dir_exists()

    if args.batch:
        # 批量处理模式
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
                # 默认输出文件名处理
                default_output = f"{theme.replace(' ', '_').replace('/', '_')}_presentation.pptx"
                output = task.get("output", default_output)

                generate_single_ppt(theme, pages, output)

        except FileNotFoundError:
            logging.error(f"批量处理文件未找到: {args.batch}")
        except json.JSONDecodeError:
            logging.error(f"解析批量处理文件JSON时出错: {args.batch}")
        except Exception as e:
            logging.error(f"批量处理过程中发生错误: {e}", exc_info=True)

    elif args.theme:
        # 单次处理模式
        logging.info("--- 开始单次生成PPT任务 ---")
        output_file = args.output if args.output else f"{args.theme.replace(' ', '_')}.pptx"
        generate_single_ppt(args.theme, args.pages, output_file)
    else:
        logging.warning("未指定操作。请使用 --theme 进行单次生成，或使用 --batch 进行批量处理。")
        parser.print_help()

if __name__ == "__main__":
    main()