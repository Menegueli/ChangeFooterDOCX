from docx import Document 

document = Document('Documento-Gabriel1.docx')  
section = document.sections[0]  
footer = section.footer  

for x in range(2,151):  
    i=x-1 

    footer.paragraphs[1].text  = footer.paragraphs[1].text.replace(f'{i}', f'{x}') 


    document.save(f'Documento-Gabriel{x}.docx')
