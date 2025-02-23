# excel_integration/excel_connector.py

import xlwings as xw
import socket
import json

def get_excel_context_data(sheet, cell_address):
    """
    Optional function to gather context from the sheet,
    e.g. text in a range of cells or the single cell's content.
    """
    return sheet.range(cell_address).value

def get_word_completions(cell_content):
    """
    Connect to the LSP server and request completions.
    """
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.connect(("localhost", 2087))  # Must match LSP server host/port
        request = json.dumps({
            "jsonrpc": "2.0",
            "id": 1,
            "method": "textDocument/completion",
            "params": {
                "text": cell_content
            }
        })
        s.sendall(request.encode())
        response = s.recv(4096)
        result = json.loads(response.decode())
        
        # Return the 'label' values from the completion items
        items = result.get("result", {}).get("items", [])
        return [item["label"] for item in items]

def excel_autocomplete():
    """
    Macro entry point for Excel to fetch suggestions for the cell A1
    and display them in B1, for example.
    """
    wb = xw.Book.caller()
    sheet = wb.sheets.active
    
    # Get context from cell A1
    user_text = get_excel_context_data(sheet, "A1")
    
    # Request completions
    suggestions = get_word_completions(user_text)
    
    # Display them in B1
    sheet.range("B1").value = ", ".join(suggestions)
