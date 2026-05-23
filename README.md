# Minha Caixinha

[![License](https://img.shields.io/github/license/guimspace/minha-caixinha)](https://github.com/guimspace/minha-caixinha/blob/master/LICENSE) [![Version](https://img.shields.io/github/release/guimspace/minha-caixinha.svg)](https://github.com/guimspace/minha-caixinha/releases)


## Overview

Minha Caixinha is a Google Sheets cash flow template. It uses Apps Script to paint the sheets and tosync upcoming bills from Google Calendar to the spreadsheet. Most of the script was derived from the larger tool [Budget n Sheets](https://github.com/budget-n-sheets/budget-n-sheets), but it is simpler in scope, and streamlined in Portuguese for Brazillians.

Given the ["Renda recorde e desemprego baixo: por que o brasileiro segue endividado, mesmo ganhando mais?"](https://g1.globo.com/economia/noticia/2026/05/08/por-que-o-brasileiro-segue-endividado.ghtml), this project is a re-use of the then cash flow template.


## Setup

### Get the template
Copy or download the spreadsheet template from one of these sources:
- Copy from [Google Drive](https://drive.google.com/file/d/1dJg7bVFnOHgztIQ4FDLENwwUGWmKRJII/view).
- Copy [Minha_Caixinha-v0.1.0.ods](templates/Minha_Caixinha-v0.1.0.ods) file.
- Download from [Internet Archive](https://archive.org/download/minhacaixinha.org-v0.1.0).

Import the template to your Google Drive. For ODS format, open it with Google Sheets.

### Create a container-bound project
1. Open the template in Google Sheets.
1. Click **Extensions** > **Apps Script**.
1. In the script editor, rename the project.
1. Copy and paste the contents of `src/main.js` and `src/index.html`.

Visit [Google Apps Script](https://developers.google.com/apps-script/overview) for additional documentation.



# License

Copyright (C) 2026 Guilherme Tadashi Maeoka

MIT License
