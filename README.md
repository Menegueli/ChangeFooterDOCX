# **Change Footer DOCX** 🤖📑

## Descrição

Esse código foi elaborado sobre a necessidade de um amigo que precisava pegar um documento já criado e replicar ele N vezes alterando o rodapé de acordo com a numeração do documento.

------

## Instalações

Para criação deste código foi necessário apenas fazer a instalação da biblioteca [Python-DOCX](https://python-docx.readthedocs.io/en/latest/) para trabalharmos com elementos dentro do Microsoft Word.

```
pip3 install python-docx
```

------

## Importação

Após a criação do seu arquivo .py devemos importar a biblioteca como abaixo.

```
from docx import Document
```

------

## Codificação

Dentro da pasta do arquivo .py adicionamos o arquivo .docx no qual queremos clonar.

Na primeira linha de código após a importação da biblioteca, definimos uma variavel que irá chamar nosso arquivo original

```
document = Document('[Nome do Arquivo].docx')
```

Logo abaixo definimos a seção que queremos editar no arquivo word, que grande parte das vezes será sempre a seção [0]

```
section = document.sections[0] 
```

Em seguida, iremos definir nessa seção o que iremos alterar, que nesse caso será o footer (rodapé)

```
footer = section.footer
```

Sendo assim iremos agora fazer um loop que pegará o arquivo original e clonar ele de acordo com o range estipulado, nesse caso, até 150 vezes.

Lembrando que o range começou do 2 pois o 1 será o arquivo original e terminando em 151 pois o range por padrão sempre para antes de chegar no numero final, logo, terminará em 150.

```
for x in range(2,151):
```

Faço uma variavel no qual ele vai pegar o valor do arquivo antigo para substituir pelo novo

```
i=x-1 
```

Aqui fazemos o replace do rodapé pegando o valor antigo (i) e dando replace para o novo (x)

```
footer.paragraphs[1].text  = footer.paragraphs[1].text.replace(f'{i}', f'{x}') 
```

Salvando o documento declarando a versão com qual número é com o (x)

```
document.save(f'[Nome do Arquivo]{x}.docx') 
```



## Código Completo

------

Dentro da pasta do arquivo .py adicionamos o arquivo .docx no qual queremos clonar.

```
from docx import Document

document = Document('[Nome do Arquivo].docx')
section = document.sections[0] 
footer = section.footer

	for x in range(2,151):
		i=x-1 
		footer.paragraphs[1].text  = footer.paragraphs[1].text.replace(f'{i}', f'{x}') 
		document.save(f'[Nome do Arquivo]{x}.docx') 
```



