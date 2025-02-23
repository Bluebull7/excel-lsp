# lsp_server/nlp_module.py
from transformers import pipeline

class NLPModule:
    def __init__(self, model_name="bert-base-uncased"):
        """
        Initialize the BERT-based pipeline or any other NLP model.
        """
        # Example: Using a fill-mask pipeline
        self.nlp_pipeline = pipeline("fill-mask", model=model_name)

    def get_context_aware_suggestions(self, context_text):
        """
        Use the BERT pipeline to generate suggestions based on context_text.
        Return a list of suggested words/phrases.
        """
        # We insert a [MASK] token as a placeholder for the predicted word(s).
        # This is a simplistic example, you can structure it to your use case.
        prompt = f"{context_text} [MASK]."
        
        # Generate predictions
        outputs = self.nlp_pipeline(prompt)
        
        # The pipeline returns a list of dictionaries with predicted tokens
        suggestions = [out["token_str"] for out in outputs]
        return suggestions
