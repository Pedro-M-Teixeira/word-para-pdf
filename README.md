[![Python 3.13](https://img.shields.io/badge/python-3.13-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

# Word → PDF (executável para Windows)

Programa com janela: escolhes uma pasta, ele converte todos os `.doc` e `.docx` para PDF
e guarda-os numa subpasta `PDF`. Corre no computador de quem o usa — nada vai para a nuvem.

## Para quem vai usar (utilizador final)

1. Descarregar `Word_para_PDF.exe` da página **Releases** do repositório.
2. Fazer duplo clique. Se o Windows mostrar “O Windows protegeu o computador”
   (o ficheiro não está assinado digitalmente): **Mais informações → Executar mesmo assim**.
3. **Escolher pasta…** → **Converter para PDF**.

Requisito: ter o **Microsoft Word** instalado (é o motor preferido) ou o **LibreOffice** (gratuito).
Não é preciso instalar Python.




## Notas

- O workflow faz um autoteste ao `.exe` (confirma que as bibliotecas foram incluídas). Se falhar,
  a Release fica sem o ficheiro e o erro aparece no registo da Actions.
- Alguns antivírus desconfiam de executáveis criados com PyInstaller (falsos positivos).
  Soluções: assinar o `.exe` (certificado pago) ou pedir à TI para o autorizar.
- Sem Word nem LibreOffice o programa avisa o que falta instalar.
- Linha de comandos: `Word_para_PDF.exe "C:\\pasta" [--subpastas] [--motor=word|libreoffice]`
  (sem janela; o resultado vê-se nos PDF criados).
