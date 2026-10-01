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

## Para quem publica (uma só vez)

1. Criar uma conta em <https://github.com> e um repositório **público** (ex.: `word-para-pdf`).
2. Carregar para o repositório os ficheiros desta pasta (**Add file → Upload files**):
   `word_para_pdf.py`, `requirements.txt`, `.gitignore` e a pasta `.github`.
   - Se a pasta `.github` não for aceite ao arrastar: **Add file → Create new file**, escrever
     `.github/workflows/build-exe.yml` como nome e colar lá o conteúdo do ficheiro.
3. **Settings → Actions → General → Workflow permissions** → **Read and write permissions** → Save.
4. **Releases → Create a new release**:
   - em *Choose a tag* escrever `v1.0.0` e escolher *Create new tag*;
   - dar um título (ex.: “Word para PDF 1.0”) e carregar em **Publish release**.
5. Abrir o separador **Actions**: o trabalho “Gerar .exe” demora cerca de 2–4 minutos.
   Quando terminar (visto verde), a Release passa a ter o ficheiro `Word_para_PDF.exe`.
6. Partilhar o link `https://github.com/<utilizador>/<repositório>/releases/latest`.

Para lançar uma versão nova: editar o código, e repetir o passo 4 com outra etiqueta (`v1.0.1`).
Para só testar sem publicar: **Actions → Gerar .exe → Run workflow**; o `.exe` fica disponível
como “artefacto” no fim da execução (descarregável durante 90 dias).

## Notas

- O workflow faz um autoteste ao `.exe` (confirma que as bibliotecas foram incluídas). Se falhar,
  a Release fica sem o ficheiro e o erro aparece no registo da Actions.
- Alguns antivírus desconfiam de executáveis criados com PyInstaller (falsos positivos).
  Soluções: assinar o `.exe` (certificado pago) ou pedir à TI para o autorizar.
- Sem Word nem LibreOffice o programa avisa o que falta instalar.
- Linha de comandos: `Word_para_PDF.exe "C:\\pasta" [--subpastas] [--motor=word|libreoffice]`
  (sem janela; o resultado vê-se nos PDF criados).
