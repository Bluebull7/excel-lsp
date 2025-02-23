# lsp_server/server.py
from pygls.server import LanguageServer
from pygls.types import (
    CompletionItem,
    CompletionList,
    CompletionParams,
    InitializeParams
)
from .nlp_module import NLPModule

class ExcelWordCompletionServer(LanguageServer):
    def __init__(self):
        super().__init__()
        self.nlp_module = NLPModule()  # Initialize BERT-based NLP

    def get_context_aware_suggestions(self, text: str):
        # Call our NLP module
        suggestions = self.nlp_module.get_context_aware_suggestions(text)
        # Convert each suggestion into an LSP CompletionItem
        return [CompletionItem(label=sugg.strip()) for sugg in suggestions]

lsp_server = ExcelWordCompletionServer()

@lsp_server.feature('textDocument/completion')
def completions(ls: ExcelWordCompletionServer, params: CompletionParams):
    # For demo, we'll just assume the 'text' is the line typed
    # (In real usage, retrieve text from the document or from user input).
    user_input_text = "User typed text"  
    suggestions = ls.get_context_aware_suggestions(user_input_text)

    return CompletionList(
        is_incomplete=False,
        items=suggestions
    )

if __name__ == "__main__":
    lsp_server.start_tcp("localhost", 2087)  # Or pass in host/port as needed
