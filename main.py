# main.py
import logging
import os
import sys
import atexit  # Import atexit for robust cleanup

# Add the parent directory to the Python path to allow imports from ppt_builder and image_service
# This is crucial for the script to find modules like ppt_builder.presentation and image_service
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from ai_service import generate_presentation_plan
from ppt_builder.presentation import PresentationBuilder
import config  # For API keys and other configurations

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', force=True)

# Global list to keep track of temporary image files for cleanup
_temp_image_files = []


def _cleanup_temp_files():
    """Deletes all temporary image files collected during the run."""
    logging.info(f"Initiating cleanup for {len(_temp_image_files)} temporary image files.")
    for file_path in list(_temp_image_files):  # Iterate over a copy to allow modification during loop
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                logging.info(f"Deleted temporary file: {file_path}")
            _temp_image_files.remove(file_path)  # Remove from list after deletion
        except OSError as e:
            logging.error(f"Error deleting temporary file {file_path}: {e}")


# Register the cleanup function to run when the program exits
atexit.register(_cleanup_temp_files)


def main():
    """
    Main function to generate a batch of PowerPoint presentations based on predefined inputs.
    """
    logging.info("Starting batch PPT generation process...")

    # Define the list of presentations to generate
    # Each item in the list is a dictionary containing parameters for one presentation
    presentations_to_generate = [
        {
            "theme": "市场营销策略",
            "num_pages": 15,
            "output_filename": "市场营销策略_演示文稿.pptx"
        },
        {
            "theme": "人工智能的未来发展",
            "num_pages": 11,
            "output_filename": "人工智能未来_演示文稿.pptx"
        },
        {
            "theme": "2024年全球经济展望",
            "num_pages": 12,
            "output_filename": "全球经济展望_演示文稿.pptx"
        }
    ]

    for i, presentation_params in enumerate(presentations_to_generate):
        theme = presentation_params["theme"]
        num_pages = presentation_params["num_pages"]
        output_filename = presentation_params["output_filename"]

        logging.info(f"\n--- Generating Presentation {i + 1}/{len(presentations_to_generate)} ---")
        logging.info(f"Theme: '{theme}', Pages: {num_pages}, Output: '{output_filename}'")

        # 1. Generate the presentation plan using AI
        logging.info(f"Requesting AI to generate plan for theme: '{theme}'...")
        presentation_plan = generate_presentation_plan(theme, num_pages)

        if not presentation_plan:
            logging.error(f"Failed to generate presentation plan for theme '{theme}'. Skipping this presentation.")
            continue  # Move to the next presentation in the batch

        logging.info("AI plan generated successfully. Starting presentation build.")

        # 2. Build the presentation using the generated plan
        try:
            # Pass the _temp_image_files list to the builder
            builder = PresentationBuilder(presentation_plan)
            # Modify SlideRenderer to append image paths to this list
            builder.slide_renderer.temp_image_files = _temp_image_files

            builder.build_presentation(output_filename)
            logging.info(f"Presentation '{output_filename}' created successfully!")
            print(f"\nPresentation '{output_filename}' created successfully!")
        except Exception as e:
            logging.error(f"An error occurred during building presentation '{output_filename}': {e}", exc_info=True)
            print(f"\nAn error occurred for '{output_filename}': {e}. Please check the logs for more details.")

    logging.info("\nBatch PPT generation process completed.")


if __name__ == "__main__":
    main()
