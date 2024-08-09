import cv2
import numpy as np
from PIL import Image
import pytesseract
from docx import Document
import os
from spellchecker import SpellChecker
import re


# Especificar la ruta completa al ejecutable de Tesseract
pytesseract.pytesseract.tesseract_cmd = r'D:/TESSERACT/tesseract.exe'

# Ruta a la carpeta que contiene las imágenes
folder_path = 'C:/Users/hp/Desktop/CAPS_TRADUS'

# Crear un nuevo documento de Word
doc = Document()

# Función para ordenar archivos numéricamente
def sort_numerically(value):
    parts = value.split('.')
    num = ''.join(filter(str.isdigit, parts[0]))
    return int(num)

# Obtener y ordenar los archivos numéricamente
files = [f for f in os.listdir(folder_path) if f.endswith(('.png', '.jpg', '.jpeg'))]
files.sort(key=sort_numerically)

# Función para preprocesar la imagen
def preprocess_image(image_path):
    img = cv2.imread(image_path, cv2.IMREAD_COLOR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    denoised = cv2.fastNlMeansDenoising(binary, None, 30, 7, 21)
    temp_filename = "temp_image.png"
    cv2.imwrite(temp_filename, denoised)
    return temp_filename

# Recorrer cada archivo en la carpeta ordenada
for filename in files:
    image_path = os.path.join(folder_path, filename)
    preprocessed_image_path = preprocess_image(image_path)
    image = Image.open(preprocessed_image_path)
    text = pytesseract.image_to_string(image, lang='eng+kor')

    # Dividir el texto por líneas
    lines = text.splitlines()

    # Inicializar una variable para almacenar el párrafo actual
    current_paragraph = []
    
    #inicializar el spellchecker
    spell = SpellChecker()

    for line in lines:
        # Si la línea no está vacía, se añade al párrafo actual
        if line.strip():
            current_paragraph.append(line.strip())
        # Si la línea está vacía y hay contenido en current_paragraph, añadirlo al documento
        elif current_paragraph:
            # Unir las líneas en un solo párrafo
            paragraph_text = " ".join(current_paragraph)
        
            # Procesar el párrafo para correcciones ortográficas
            corrected_paragraph = re.sub(r'\d', '', paragraph_text)
            corrected_paragraph2 = " ".join([spell.correction(word) if spell.correction(word) else word for word in corrected_paragraph.split()])

        
        
            # Añadir el párrafo corregido al documento
            doc.add_paragraph(corrected_paragraph2)
        
            # Limpiar el current_paragraph para el siguiente bloque
            current_paragraph = []

# Añadir el último párrafo si existe
if current_paragraph:
    paragraph_text = " ".join(current_paragraph)
    corrected_paragraph = re.sub(r'\d', '', paragraph_text)
    corrected_paragraph2 = " ".join([spell.correction(word) if spell.correction(word) else word for word in corrected_paragraph.split()])

    doc.add_paragraph(corrected_paragraph2)

# Eliminar la imagen temporal
os.remove(preprocessed_image_path)

# Especificar la ruta completa donde se guardará el documento
output_path = 'C:/Users/hp/Desktop/TEXTO.docx'

# Guardar el documento
doc.save(output_path)