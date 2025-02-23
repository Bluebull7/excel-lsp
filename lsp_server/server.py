# lsp_server/server.py

from pygls.server import LanguageServer
from lsprotocol.types import (
    CompletionItem,
    CompletionList,
    CompletionParams,
    InitializeParams
)
from loguru import logger

from nlp_module import NLPModule

# Configure loguru (Customize as needed)
logger.add("logs/server.log", rotation="1 MB", level="INFO")

class ExcelWordCompletionServer(LanguageServer):
    def __init__(self):
        super().__init__()
        
        try:
            self.nlp_module = NLPModule()
            logger.info("NLPModule successfully initialized.")
        except Exception as e:
            logger.exception("Failed to initialize NLPModule.")
            # Depending on your use case, you might want to re-raise 
            # or exit the program here if the module is critical.
            raise e

    def get_context_aware_suggestions(self, text: str):
        """
        Retrieves context-aware suggestions from the NLP module.
        Logs detailed information for debugging and error handling.
        """
        logger.debug(f"Fetching context-aware suggestions for text: {text}")
        try:
            suggestions = self.nlp_module.get_context_aware_suggestions(text)
            logger.debug(f"Suggestions returned: {suggestions}")
            return [CompletionItem(label=sugg.strip()) for sugg in suggestions]
        except Exception as e:
            logger.exception("Error while generating context-aware suggestions.")
            # Return an empty list or a fallback suggestion
            return []

lsp_server = ExcelWordCompletionServer()

@lsp_server.feature("textDocument/completion")
def completions(ls: ExcelWordCompletionServer, params: CompletionParams):
    """
    Handle a 'textDocument/completion' request using the BERT-based context suggestions.
    """
    try:
        # For demonstration, let's assume the user text is accessible via params
        # or you might retrieve it from the buffer if you have a text document open.
        user_input_text = params.text if hasattr(params, 'text') else "User typed text"

        logger.info(f"Received completion request for text: {user_input_text}")
        suggestions = ls.get_context_aware_suggestions(user_input_text)
        
        completion_list = CompletionList(
            is_incomplete=False,
            items=suggestions
        )

        logger.info(f"Returning {len(suggestions)} suggestions to the client.")
        return completion_list

    except Exception as e:
        logger.exception("Unhandled error in 'completions' handler.")
        # Return an empty CompletionList or handle the error as needed
        return CompletionList(is_incomplete=False, items=[])

if __name__ == "__main__":
    logger.info("Starting ExcelWordCompletionServer on localhost:2087...")
    try:
        lsp_server.start_tcp("localhost", 2087)
    except Exception as e:
        logger.exception("Failed to start LSP server.")
        raise e
