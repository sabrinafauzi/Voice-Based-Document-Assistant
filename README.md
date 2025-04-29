# Voice-Based-Document-Assistant

This assistant is designed to help visually impaired users read and understand documents in PDF and DOCX formats through voice commands. It combines Speech-to-Text, Text-to-Speech, and local Large Language Model (LLM) technologies to create a natural and accessible user experience.

# Main Features:
* Voice commands for navigating documents.
* Support for PDF and DOCX file formats.
* Automatic summarization of document content.
* Question answering based on document content.
* Voice responses using TTS.
* Runs locally without needing an internet connection.
* Keyboard input (e.g., f, j, space, esc) for navigation, validation, and interruption.

# Technologies Used:
* Speech-to-Text (STT): Vosk
* Text-to-Speech (TTS): pyttsx3
* NLP/LLM: Local Mistral model via LM Studio
* Document reader: python-docx, PyMuPDF (fitz)
