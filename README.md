# Chainsaw 🪚

**Sistema de Padronização de Proposituras Legislativas**

Chainsaw is an advanced, robust VBA macro project designed exclusively for Microsoft Word. It automatically sanitizes, structures, and formats complex Brazilian legislative documents (Proposituras), turning raw, unformatted text into perfectly aligned, standardized legal documents.

## ✨ Features

- **Automated Element Identification:** Uses an advanced heuristic engine to detect structural parts of a document such as *Título*, *Ementa*, *Justificativa*, *Data (Plenário)*, and *Assinaturas*.
- **Two-Pass Formatting Pipeline:** 
  - **Pass 1:** Normalizes syntax, limits blank lines, clears arbitrary text boundaries, and purges bad encodings.
  - **Pass 2:** Applies pixel-perfect layout alignments, margins, and custom indentation rules seamlessly.
- **Visual Content Protection:** Explicitly caches and protects images, bullet points, numbered lists, and header layouts from getting destroyed during the document formatting process.
- **Intelligent Clause Parsing:** Finds and formats specific legal clauses heavily used in Brazilian legislation like *“Considerando”*, *“Ante o Exposto”*, and *“In Loco”*.
- **Failsafe Executions:** Includes heavily engineered error recovery, undo-group wraps (`UndoRecord`), and pre-execution backups to ensure MS Word never critically fails during processing.

## 🏗️ Architecture

The project has been scaled into 7 robust, interoperable VBA modules located inside the `source/main/` directory:

- `ModConfig.bas`: Central system configs, user-defined state structures (`UDTs`), and UI view states.
- `ModCore.bas`: Cache pipelines, structural heuristics, and validation checks.
- `ModMain.bas`: Macro entry points (`PadronizarDocumentoMain`, `concluir`) exposed to the Word Ribbon/UI.
- `ModProcess.bas`: The main engine applying text formats, title alignments, space normalizations, and explicit character formats.
- `ModMedia.bas`: Image tracking, indentation overrides for bulleted lists, and media caching mechanisms.
- `ModSystem.bas`: Telemetry mapping, OS-level progress bars, fast backups, and unhandled-exception recovery logic.
- `ModUtils.bas`: Stateless path logic, general IO, and strictly safe COM object wrappers.

## 🚀 Installation & Usage

1. Open Microsoft Word.
2. Launch the **Visual Basic for Applications (VBA) Editor** (`ALT` + `F11`).
3. Import the 7 `.bas` files found in the `source/main/` folder into your `Normal.dotm` or dedicated Document Template.
4. Go to `Debug -> Compile Project` to ensure your Word environment resolves the inter-module Public references.
5. Create a Ribbon Button or Quick Access Toolbar shortcut pointing to the `PadronizarDocumentoMain` macro.
6. Click the macro while editing a document to execute the Chainsaw standardized pipeline!

## 📜 License

This project is licensed under the **GNU GPLv3** License. See the `LICENSE` file for more details.
