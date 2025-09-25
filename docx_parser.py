import os
import json
import zipfile
from lxml import etree


class DocxParser:
    def __init__(self):
        # Namespace para documentos Word
        self.namespaces = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        }
        self.segments = []

    def extract_text_from_docx(self, docx_path):
        """
        Extrae texto de un archivo DOCX usando técnica Roundtripping
        """
        print(f"Procesando archivo: {docx_path}")

        try:
            # se descomprime el archivo DOCX
            with zipfile.ZipFile(docx_path, "r") as zip_file:
                # listamos contenido zip para debug
                print("Archivos en el DOCX:")
                for name in zip_file.namelist():
                    print(f"  - {name}")

                # leemos el document.xml principal
                if "word/document.xml" in zip_file.namelist():
                    xml_content = zip_file.read("word/document.xml")
                else:
                    raise FileNotFoundError(
                        "No se encontró word/document.xml en el archivo DOCX"
                    )

            # se parsea el XML con lxml
            root = etree.fromstring(xml_content)

            # extraer párrafos (w:p) y texto (w:t)
            paragraphs = root.xpath("//w:p", namespaces=self.namespaces)
            print(f"Párrafos encontrados: {len(paragraphs)}")

            segment_counter = 1

            for paragraph in paragraphs:
                # se extraen todos los elementos de texto (w:t) dentro del párrafo
                text_elements = paragraph.xpath(".//w:t", namespaces=self.namespaces)

                if text_elements:
                    # concatenar todo el texto del párrafo
                    paragraph_text = "".join(
                        [elem.text or "" for elem in text_elements]
                    )
                    paragraph_text = paragraph_text.strip()

                    # solo agregar párrafos que contengan texto
                    if paragraph_text:
                        segment = {
                            "id": f"seg-{segment_counter:04d}",
                            "key": f"paragraph-{segment_counter}",
                            "source": paragraph_text,
                            "target": "",
                        }
                        self.segments.append(segment)
                        segment_counter += 1
                        print(f"Párrafo extraído: {paragraph_text[:50]}...")

            print(f"Total de segmentos extraídos: {len(self.segments)}")

        except Exception as e:
            print(f"Error al procesar el archivo DOCX: {e}")
            raise

    def save_json(self, output_path):
        """
        Guarda los segmentos en formato JSON con indentación nivel 4 y encoding UTF-8
        """
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(self.segments, f, indent=4, ensure_ascii=False)
            print(f"Archivo JSON guardado: {output_path}")
            print(f"Segmentos guardados: {len(self.segments)}")

        except Exception as e:
            print(f"Error al guardar el archivo JSON: {e}")
            raise

    def process_docx_file(self, docx_path, json_output_path):
        """
        Proceso completo: extrae texto de DOCX y guarda JSON
        """
        self.segments = []  # Reiniciar la lista
        self.extract_text_from_docx(docx_path)
        self.save_json(json_output_path)

        return len(self.segments)


def setup_directories():
    # creación de directorio
    dirs = [
        "Python",
        "Python/Parser",
        "Python/Parser/DOCX",
        "Python/Parser/DOCX/sample",
    ]

    for directory in dirs:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"Directorio creado: {directory}")
        else:
            print(f"Directorio ya existe: {directory}")


def main():

    print("=== DOCX Parser - Técnica Roundtripping ===\n")

    # creación de directorios
    setup_directories()

    # se crea la instancia de la clase
    parser = DocxParser()

    # se buscan los archivos DOCX en la carpeta sample
    samples_dir = "Python/Parser/DOCX/sample"
    # listamos y buscamos los archivos que terminen en .docx
    docx_files = [f for f in os.listdir(samples_dir) if f.endswith(".docx")]

    if not docx_files:
        print(f"No se encontraron archivos .docx en {samples_dir}")
        print(
            "Por favor, coloca un archivo .docx en esa carpeta y ejecuta el script nuevamente."
        )

        return

    # Procesar cada archivo DOCX encontrado
    for docx_file in docx_files:
        print(f"\n--- Procesando: {docx_file} ---")

        docx_path = os.path.join(samples_dir, docx_file)
        json_filename = f"{os.path.splitext(docx_file)[0]}_segments.json"
        json_path = os.path.join("Python/Parser/DOCX", json_filename)

        try:
            #utilizo método de la clase para procesar el archivo
            segments_count = parser.process_docx_file(docx_path, json_path) 
            print(f"Procesado exitosamente: {segments_count} segmentos extraídos")
            print(f"Archivo JSON generado: {os.path.abspath(json_path)}")

            # una muestra de los primeros segmentos extraídos
            if parser.segments:
                print("\n--- Muestra de segmentos extraídos ---")
                for i, segment in enumerate(
                    parser.segments[:3]
                ):  # Mostrar solo los primeros 3
                    print(f"{segment['id']}: {segment['source'][:80]}...")
                if len(parser.segments) > 3:
                    print(f"... y {len(parser.segments) - 3} segmentos más")

        except Exception as e:
            print(f"Error al procesar {docx_file}: {e}")


if __name__ == "__main__":
    main()
