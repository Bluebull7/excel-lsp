# 📊 Excel Word Completion & Autocorrect LSP Server

**A local Python-based Language Server Protocol (LSP) server for intelligent word completion and autocorrect in Excel, integrated using `xlwings`.** This project provides real-time word suggestions and autocorrect functionalities directly in Excel cells, with a roadmap for advanced NLP and UI features.

---

## 🚀 Features

- ✅ **Word Completion** — Intelligent suggestions for partial words typed in Excel.
- ✅ **Autocorrect** — Basic autocorrect functionality using `textblob`.
- ✅ **Context-Aware NLP** — Advanced completions using BERT for smarter suggestions.
- ✅ **Local LSP Server** — Lightweight and fast, built using `pygls`.
- ✅ **Excel Integration** — Seamlessly connects Excel to the LSP server via `xlwings`.
- ✅ **Modular Architecture** — Optimized design patterns for improved modularity and structure.
- ✅ **Easy Setup** — Minimal dependencies and simple architecture.

---

## 📁 Project Structure

```
📦 excel-lsp-autocomplete
├── lsp_server/
│   ├── __init__.py          # LSP server package
│   ├── server.py            # Main LSP server logic
│   └── nlp_module.py        # Context-aware NLP logic using BERT
├── excel_integration/
│   ├── __init__.py          # Excel integration package
│   └── excel_connector.py   # xlwings integration with Excel
├── requirements.txt        # Python dependencies
├── README.md               # This file
└── examples/
    └── example.xlsx        # Example Excel file with macros
```

---

## 🛠️ Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/your-username/excel-lsp-autocomplete.git
   cd excel-lsp-autocomplete
   ```

2. **Set up a virtual environment (optional but recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Run the LSP server:**
   ```bash
   python -m lsp_server.server
   ```

5. **Open the Excel example file:**
   - Navigate to the `examples/` folder.
   - Open `example.xlsx` with macros enabled.
   - Run the macro to test word completion.

---

## 📋 Usage

1. **Type in cell A1** in the example Excel file.
2. **Run the macro** (via `xlwings`) to trigger word completion.
3. Suggested completions will appear in **cell B1**.
4. Enjoy context-aware suggestions powered by BERT! 🧠

---

## ⚡ Technologies Used

- [Python](https://www.python.org/)
- [pygls](https://pypi.org/project/pygls/) — Python Language Server framework.
- [xlwings](https://www.xlwings.org/) — Python for Excel integration.
- [textblob](https://textblob.readthedocs.io/en/dev/) — Simple text processing for autocorrect.
- [transformers](https://huggingface.co/docs/transformers/) — BERT integration for context-aware NLP.

---

## 🔮 Future Enhancements

- ⚡ **Real-time Suggestions**:
  - Implement WebSocket-based communication for live word suggestions as users type.

- 🛡️ **Enhanced Autocorrect**:
  - Integrate `symspellpy` for high-speed autocorrect with large dictionaries.

- 🎨 **Improved UI in Excel**:
  - Add dropdowns or suggestion lists directly in Excel cells using VBA forms or Data Validation.

- 📈 **Deeper Contextual Awareness**:
  - Expand NLP logic to analyze larger cell ranges and cross-sheet data.

---

## 🤝 Contributing

Contributions are welcome! Please open an issue or submit a pull request.

1. Fork the repo
2. Create your feature branch (`git checkout -b feature/YourFeature`)
3. Commit your changes (`git commit -m 'Add some feature'`)
4. Push to the branch (`git push origin feature/YourFeature`)
5. Open a Pull Request

---

## 📄 License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## 🌟 Acknowledgements

- Inspired by the flexibility of **LSP** and the power of **Python for Excel**.
- Special thanks to open-source communities behind `pygls`, `xlwings`, `textblob`, and `transformers`.

---

