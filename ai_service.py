# ai_service.py
import json
import logging
from openai import OpenAI
import config

# --- Initialize Client ---
try:
    client = OpenAI(
        api_key=config.ONEAPI_KEY,
        base_url=config.ONEAPI_BASE_URL,
    )
    logging.info("OpenAI client initialized.")
except Exception as e:
    logging.error(f"Failed to initialize OpenAI client: {e}")
    client = None

def load_prompt_template(file_path: str = "prompt_template.md") -> str | None:
    """
    Loads the prompt template from a file.

    Args:
        file_path (str): The path to the prompt template file.

    Returns:
        str | None: The content of the template file, or None if an error occurs.
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        logging.error(f"Prompt template file not found at: {file_path}")
        return None
    except Exception as e:
        logging.error(f"Error reading prompt template file: {e}", exc_info=True)
        return None

def _extract_json_from_response(text: str) -> str | None:
    """
    Extracts a JSON object from a string that may contain additional text.
    """
    try:
        # Find the first '{' and the last '}'
        start_index = text.find('{')
        end_index = text.rfind('}')
        if start_index != -1 and end_index != -1 and end_index > start_index:
            return text[start_index:end_index+1]
        logging.warning("Could not find a valid JSON object in the AI response.")
        return None
    except Exception as e:
        logging.error(f"Error while extracting JSON: {e}")
        return None


def generate_presentation_plan(theme: str, num_pages: int) -> dict | None:
    """
    Generates a detailed JSON plan for a presentation using the OpenAI API.
    This version loads the prompt from an external template file.

    Args:
        theme (str): The theme of the presentation.
        num_pages (int): The desired number of pages.

    Returns:
        dict | None: A dictionary containing the presentation plan, or None on failure.
    """
    if not client:
        logging.error("OpenAI client not initialized. Cannot generate presentation plan.")
        return None

    # Load the prompt template from the file
    prompt_template = load_prompt_template()
    if not prompt_template:
        return None

    # Format the prompt with the dynamic variables
    prompt = prompt_template.format(theme=theme, num_pages=num_pages)
    try:
        logging.info("Generating presentation plan from AI...")
        response = client.chat.completions.create(
            model=config.MODEL_NAME,
            messages=[
                {"role": "system", "content": "You are a world-class presentation designer. Your output must be a single, raw JSON object without any extra text or markdown. You must strictly follow all instructions."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.6,
        )

        response_content = response.choices[0].message.content
        if response_content:
            logging.info("Successfully received presentation plan from AI.")
            json_string = _extract_json_from_response(response_content)
            if json_string:
                return json.loads(json_string)
            else:
                logging.error("Could not extract a valid JSON object from the AI's response.")
                logging.debug(f"Raw response was: {response_content}")
                return None
        else:
            logging.error("Received empty response from AI.")
            return None

    except json.JSONDecodeError as e:
        logging.error(f"Failed to parse JSON response from AI: {e}")
        logging.error(f"Raw response was: {response_content}")
        return None
    except Exception as e:
        logging.error(f"An error occurred while communicating with OpenAI: {e}", exc_info=True)
        return None
