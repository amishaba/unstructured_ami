import logging
import requests
from fastmcp.tools import Tool
from fastmcp.utilities.func_metadata import ArgModelBase
from prototype.shared_libraries.config_mediator import check_if_user_access

# Initialize logging
logger = logging.getLogger("pptx_structure_tool")
logging.basicConfig(level=logging.INFO)


class PPTXStructureArgs(ArgModelBase):
    pptx_file_path: str  # Full path to the PPTX file to be processed


class PPTXStructureTool(Tool):
    def __init__(self):
        """
        Initialize the PPTXStructureTool with metadata and runtime parameters.
        """
        logger.debug("Initializing PPTXStructureTool")
        super().__init__(
            name="PPTXStructureTool",
            description="Extract structured and categorized elements from a PPTX file using a microservice.",
            fn=self.run_pptx_structure_tool,
            parameters={
                "pptx_file_path": {
                    "type": "string",
                    "description": "Full path to the PPTX file to extract content from.",
                }
            },
            fn_metadata={
                "name": "run_pptx_structure_tool",
                "description": "Calls the microservice to extract structure from PPTX.",
                "arg_model": PPTXStructureArgs,
            },
            is_async=False,
        )

    @classmethod
    @check_if_user_access
    def run_pptx_structure_tool(cls, pptx_file_path: str, config: dict = None) -> dict:
        """
        Sends the PPTX file to a microservice endpoint for structured extraction.

        Args:
            pptx_file_path (str): Local filesystem path to the PPTX file.
            config (dict): Dictionary containing microservice endpoint settings.

        Returns:
            dict: Structured output from microservice, or error information.
        """
        logger.info(f"Attempting to extract structure from file: {pptx_file_path}")

        # Check config for required URL
        if not config or "pptx_structure_microservice_url" not in config:
            logger.error("Missing 'pptx_structure_microservice_url' in config")
            return {"error": "Configuration missing microservice URL."}

        try:
            # Prepare the file payload for upload
            with open(pptx_file_path, "rb") as pptx_file:
                file_payload = {
                    'file': (
                        pptx_file_path,
                        pptx_file,
                        'application/vnd.openxmlformats-officedocument.presentationml.presentation'
                    )
                }

                # Send POST request to microservice
                response = requests.post(
                    url=config["pptx_structure_microservice_url"],
                    files=file_payload,
                    timeout=30
                )

            response.raise_for_status()
            logger.info(f"Received response from microservice for file: {pptx_file_path}")
            return response.json()

        except FileNotFoundError:
            logger.exception("File not found on disk")
            return {"error": f"File not found: {pptx_file_path}"}

        except requests.exceptions.RequestException as request_error:
            logger.exception("Request to microservice failed")
            return {"error": f"Request failed: {str(request_error)}"}

        except Exception as unexpected_error:
            logger.exception("Unexpected error during PPTX structure extraction")
            return {"error": f"Unexpected error: {str(unexpected_error)}"}
