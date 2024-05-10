from django.shortcuts import render # type: ignore
from django.shortcuts import render # type: ignore
from django.http import HttpResponse # type: ignore
from django.core.files.storage import FileSystemStorage # type: ignore
import tabula # type: ignore
import pandas as pd # type: ignore
import os
from django.core.files.storage import FileSystemStorage # type: ignore
from pdf2docx import Converter # type: ignore

def homepage(request):
    return render(request, "html/index.html")

# def pdftoexcel(request):
#     if request.method == 'POST' and request.FILES.get('pdf_file'):
#         # Récupérer le fichier PDF depuis la requête
#         pdf_file = request.FILES['pdf_file']
        
#         # Enregistrer le fichier PDF dans le système de fichiers temporaire de Django
#         fs = FileSystemStorage()
#         filename = fs.save(pdf_file.name, pdf_file)
#         pdf_file_path = fs.path(filename)
        
#         try:
#             # Extraction des données du PDF
#             df = tabula.read_pdf(pdf_file_path, pages='all')
            
#             # Fusionner toutes les données extraites en un seul DataFrame
#             merged_df = pd.concat(df)
            
#             # Création du chemin pour le fichier Excel
#             excel_file_path = os.path.splitext(pdf_file_path)[0] + '.xlsx'
            
#             # Écriture du DataFrame dans un fichier Excel
#             merged_df.to_excel(excel_file_path, index=False)
            
#             # Téléchargement du fichier Excel
#             with open(excel_file_path, 'rb') as excel_file:
#                 response = HttpResponse(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
#                 response['Content-Disposition'] = 'attachment; filename="' + os.path.basename(excel_file_path) + '"'
#                 return response
#         except Exception as e:
#             return HttpResponse("Erreur lors de la conversion : {}".format(str(e)))
#     return render(request, 'html/pdftoexcel.html')
def pdftoexcel(request):
    if request.method == 'POST' and request.FILES.get('pdf_file'):
        pdf_file = request.FILES['pdf_file']
        fs = FileSystemStorage()
        filename = fs.save(pdf_file.name, pdf_file)
        pdf_file_path = fs.path(filename)
        
        try:
            # Extraction des données du PDF
            df = tabula.read_pdf(pdf_file_path, pages='all')
            
            # Fusionner toutes les données extraites en un seul DataFrame
            merged_df = pd.concat(df)
            
            # Création du chemin pour le fichier Excel
            excel_file_path = os.path.splitext(pdf_file_path)[0] + '.xlsx'
            
            # Écriture du DataFrame dans un fichier Excel
            merged_df.to_excel(excel_file_path, index=False)
            
            # Transmettre le chemin du fichier Excel au template
            context = {'excel_file_path': excel_file_path}
            
            # Rendre le template HTML avec le contexte
            return render(request, 'html/pdftoexcel.html', context)
        except Exception as e:
            return HttpResponse("Erreur lors de la conversion : {}".format(str(e)))
    return render(request, 'html/pdftoexcel.html')

def pdftoword(request):
    if request.method == 'POST' and request.FILES.get('pdf_file'):
        pdf_file = request.FILES['pdf_file']
        
        # Enregistrer le fichier PDF dans le système de fichiers temporaire de Django
        fs = FileSystemStorage()
        filename = fs.save(pdf_file.name, pdf_file)
        pdf_file_path = fs.path(filename)
        
        try:
            # Convertir le PDF en Word
            word_file_path = pdf_file_path.replace('.pdf', '.docx')
            cv = Converter(pdf_file_path)
            cv.convert(word_file_path, start=0, end=None)
            cv.close()
            
            # Téléchargement du fichier Word
            with open(word_file_path, 'rb') as word_file:
                response = HttpResponse(word_file.read(), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                response['Content-Disposition'] = 'attachment; filename="' + pdf_file.name.replace('.pdf', '.docx') + '"'
                return response
        except Exception as e:
            return HttpResponse("Erreur lors de la conversion : {}".format(str(e)))
    
    return render (request, 'html/pdftoword.html')