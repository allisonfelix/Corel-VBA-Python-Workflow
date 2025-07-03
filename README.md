# CorelDRAW Batch Automation with Python + VBA
### <i> Automação em Lote no Corel com Python + VBA</i>

Este repositório demonstra como automatizar, em lote, a execução de macros VBA (GMS) no CorelDRAW a partir de um script Python, usando COM (pywin32). O código percorre pastas raiz, identifica arquivos de arte (JPG, PNG, TIF e CDR), executa verificações (fontes faltantes, associação a PDFs) e dispara macros em sequência, sem interação manual.

> **Atenção**: o arquivo **Graficonauta.gms** não está disponível aqui, pois contém rotinas desenvolvidas para uso interno de uma empresa. Este projeto, contudo, exemplifica as estratégias de integração Python ↔ CorelDRAW/VBA e pode ser adaptado para qualquer conjunto de macros.

## Índice

- [Requisitos](#requisitos)  
- [Configuração](#configuração)  
- [Uso](#uso)  
- [Como funciona](#como-funciona)  
- [Personalização](#personalização)  
- [Licença](#licença)  

## Requisitos

- Windows 10/11  
- Python 3.8+  
- `pywin32`  
- CorelDRAW X7 ou superior (testado até CorelDRAW 2024)  
- MuPDF/`mutool` ou GhostScript  

Instale as dependências Python:

```bash
pip install pywin32
```

## Configuração

Abra o arquivo `tratar-python.py` e ajuste:

- **ROOT_DIRS**: pastas raiz para varredura.  
- **ALLOWED_EXT**: extensões permitidas.  
- **MAX_FILES_PER_SUBFOLDER**: limite de arquivos por subpasta.  
- **GMS_PROJECT**, **GMS_MODULE**, **GMS_PROCEDURE**: nomes do seu `.gms`, módulo e procedimento.  
- **GMS_REFUGO_POR_FONTE**: macro para arquivos com fontes faltantes.  
- **IGNORED_SYSTEM_FONTS**: fontes do sistema a ignorar.  
- **ONLY_DIGITAL_CDR** + **CDR_KEYWORDS**: filtro de CDRs (ex.: “impressao-digital”).  
- **AUTO_CLOSE_MULTIPLE**: fecha documentos extras.  
- **POLL_TIMEOUT**: tempo máximo de espera por macro.

## Uso

1. Posicione seu GMS (por ex. `Graficonauta.gms`) em `%ProgramFiles%\Corel\CorelDRAW Graphics Suite <versão>\Draw\GMS`, onde `<versão>` é o ano da sua instalação (por ex. `2022` ou `2024`).  
2. Ajuste configurações em `tratar-python.py`.  
3. Execute:

   ```bash
   python tratar-python.py
   ```

4. O script roda em loop, processando cada arquivo e exibindo logs no console.

## Como funciona

1. **COM**: conecta/reconecta ao CorelDRAW via `pywin32`.  
2. **Varredura**: percorre `ROOT_DIRS`, ignora subpastas específicas.  
3. **Filtragem**: identifica arquivos por extensão e palavras-chave.  
4. **Fontes**: checa `MissingFontList` (fallback `Fonts`), ignora Arial/Calibri.  
5. **PDF**: para TIF “impressao-digital”, busca PDF e conta páginas.  
6. **Macros**: executa `GMSManager.RunMacro` e faz polling com `app.Busy` ou `Pages.Count`.  
7. **Resiliência**: reabre Corel se cair, faz `gc.collect()`, loop estável.

## Personalização

- Adicione/remova **CDR_KEYWORDS**.  
- Inclua novos formatos em **ALLOWED_EXT**.  
- Ajuste método de polling (e.g. `app.GMSManager.IsBusy`).  
- Extenda para outros cenários COM ou formatos de arquivo.

## Licença

Mozilla Public License Version 2.0
