import logging
import argparse
import json
from ai_service import generate_presentation_plan
from ppt_builder.presentation import PresentationBuilder

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',force=True)


def generate_single_ppt(theme: str, num_pages: int, output_file: str):
    """
    为单个主题生成演示文稿。
    """
    logging.info(f"Theme: '{theme}', Pages: {num_pages}, Output: '{output_file}'")

    # 1. 从AI获取演示文稿计划
    logging.info(f"Requesting AI to generate plan for theme: '{theme}'...")
    plan = generate_presentation_plan(theme, num_pages)

    if plan:
        logging.info("AI plan generated successfully. Starting presentation build.")
        # 2. 根据计划构建演示文稿
        try:
            builder = PresentationBuilder(plan)
            builder.build_presentation(output_file)
            logging.info(f"Presentation generation complete. Saved to {output_file}")
            return True
        except Exception as e:
            logging.error(f"Failed to build presentation for theme '{theme}': {e}", exc_info=True)
            return False
    else:
        logging.error(f"Failed to generate a presentation plan for theme '{theme}'. Aborting this task.")
        return False


def main():
    """
    主函数，支持通过命令行参数进行单次生成，或通过配置文件进行批量生成。
    """
    parser = argparse.ArgumentParser(description="AI PPT Generator")
    parser.add_argument("--theme", type=str, help="The theme of the presentation (for single mode).")
    parser.add_argument("--pages", type=int, default=10, help="The number of pages for the presentation.")
    parser.add_argument("--output", type=str, help="The output file name for the presentation (for single mode).")
    parser.add_argument("--batch", type=str, help="Path to a JSON file for batch processing.")

    args = parser.parse_args()

    if args.batch:
        # 批量处理模式
        logging.info("Starting batch PPT generation process...")
        try:
            with open(args.batch, 'r', encoding='utf-8') as f:
                batch_tasks = json.load(f)

            total_tasks = len(batch_tasks)
            for i, task in enumerate(batch_tasks):
                logging.info(f"\n--- Generating Presentation {i + 1}/{total_tasks} ---")
                theme = task.get("theme")
                pages = task.get("pages", args.pages)
                output = task.get("output", f"{theme.replace(' ', '_')}_presentation.pptx")

                if not theme:
                    logging.warning(f"Skipping task {i + 1} due to missing theme.")
                    continue

                generate_single_ppt(theme, pages, output)

        except FileNotFoundError:
            logging.error(f"Batch file not found: {args.batch}")
        except json.JSONDecodeError:
            logging.error(f"Error decoding JSON from batch file: {args.batch}")
        except Exception as e:
            logging.error(f"An error occurred during batch processing: {e}", exc_info=True)

    elif args.theme:
        # 单次处理模式
        logging.info("Starting single PPT generation process...")
        output_file = args.output if args.output else f"{args.theme.replace(' ', '_')}.pptx"
        generate_single_ppt(args.theme, args.pages, output_file)
    else:
        logging.warning("No action specified. Use --theme for a single PPT or --batch for batch processing.")
        parser.print_help()


if __name__ == "__main__":
    main()
