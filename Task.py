from docx import Document
from docx.oxml.ns import qn
from transformers import MarianMTModel, MarianTokenizer
import torch


model_name = 'Helsinki-NLP/opus-mt-en-es'
tokenizer = MarianTokenizer.from_pretrained(model_name)
model = MarianMTModel.from_pretrained(model_name)

def translate_text(text, tokenizer, model):
    if not text.strip():  
        return text
    inputs = tokenizer(text, return_tensors="pt", padding=True, truncation=True)
    with torch.no_grad():
        translated = model.generate(**inputs)
    translated_text = tokenizer.batch_decode(translated, skip_special_tokens=True)
    return translated_text[0]

def translate_run(run, tokenizer, model):
    english_text = run.text
    if english_text.strip(): 
        try:
            spanish_text = translate_text(english_text, tokenizer, model)
            run.text = spanish_text
        except Exception as e:
            print(f"Error translating run: {english_text}\nError: {e}")
            run.text = english_text  

def translate_docx(input_file, output_file, tokenizer, model):
    doc = Document(input_file)
    for para in doc.paragraphs:
        for run in para.runs:
            translate_run(run, tokenizer, model)
            
    doc.save(output_file)

# File paths
input_file = 'input.docx'
output_file = 'output.docx'

# Translate the document
translate_docx(input_file, output_file, tokenizer, model)
