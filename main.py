# main.py
import os
import logging
import config
from ai_service import generate_presentation_plan
from ppt_builder.presentation import build_presentation

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', force=True)


def process_theme(theme: str):
    """Generates a presentation for a single theme."""
    logging.info(f"--- Starting process for theme: '{theme}' ---")

    # Pass the number of pages from config to the plan generator
    ppt_plan = generate_presentation_plan(theme, config.NUM_PAGES)

    if ppt_plan:
        safe_theme_name = "".join(c for c in theme if c.isalnum() or c in " _-").rstrip()
        output_filename = f"Generated_{safe_theme_name}.pptx"
        output_filepath = os.path.join(config.OUTPUT_DIR, output_filename)

        build_presentation(ppt_plan, output_filepath)
    else:
        logging.error(f"Failed to generate data for theme '{theme}'. Skipping.")

    logging.info("-" * 50)


def main():
    """Main function to run the PPT generation process."""
    if "YOUR_ONEAPI_KEY_HERE" in config.ONEAPI_KEY or "YOUR_PEXELS_API_KEY_HERE" in config.PEXELS_API_KEY:
        logging.error("Please set your API keys in a .env file or as environment variables in config.py.")
        return

    if not os.path.exists(config.OUTPUT_DIR):
        os.makedirs(config.OUTPUT_DIR)
        logging.info(f"Created output directory: {config.OUTPUT_DIR}")

    themes = [
        "企业数字化转型解决方案",
        "新消费品牌市场进入策略",
        "人工智能在教育领域的应用"
    ]

    for theme in themes:
        process_theme(theme)


if __name__ == "__main__":
    main()