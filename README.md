# Net-zero implementation among transport-sector SMEs in England

This repository contains the survey data, qualitative coding and Python analysis supporting the research paper:

**Ahmadinia, M., Spooner, S., Khasraw, D., Long, T. and Kauffman, M.** *Net Zero, Not Zero Obstacles: Understanding Decarbonisation Challenges Among Transport SMEs in England* Publication details and DOI will be added when available.

The study analyses responses from **44 SMEs in England** using 27 five-point questionnaire statements and five open-ended questions.

## Repository files

| File | Description |
| --- | --- |
| `Analysis.xlsx` | Survey responses, analysis worksheets and validation checks |
| `qualitative_coding.csv` | Coding of open-text responses |
| `scope_classification.csv` | Transport-activity classification |
| `reproduce.py` | Python code to reproduce analyses, tables and figures |
| `requirements.txt` | Python dependencies |
| `README.md` | Repository information and instructions |

## Run the analysis

Keep the files in the same folder. Run:

```bash
python -m pip install -r requirements.txt
python reproduce.py
```

The script creates `outputs/` and `figures/` folders. A successful run ends with **`VALIDATION PASSED`**. The tested workflow reproduced the reported results with **0 mismatches across 992 manuscript-value checks and 955 spreadsheet comparisons**. The script generates the original conceptual Figure 3, which may differ visually from the final journal figure.

## Citation

**If you use, adapt or build on this code, dataset, methods or figures in academic work, please cite the associated research paper and acknowledge this repository.** Update the reference with the published journal details and DOI when available.

Suggested repository citation:

> Ahmadinia, M., Spooner, S., Khasraw, D., Long, T. and Kauffman, M. (2026). *Net-zero implementation among transport-sector SMEs in England: data, coding and reproducibility materials*. GitHub repository.

## Licence

**All materials in this repository, including the Python code, datasets, figures and documentation, are released under the Creative Commons Attribution 4.0 International (CC BY 4.0) Licence.**

You may copy, use, share and adapt these materials, including for commercial purposes, provided you give appropriate credit, link to the licence and indicate any changes. Attribution must not suggest endorsement by the authors.

Full licence: <https://creativecommons.org/licenses/by/4.0/>

## Contact

**Dr Masoud Ahmadinia**  
Institute for Advanced Manufacturing and Engineering (AME), Coventry University, United Kingdom  
**ae4045@coventry.ac.uk**

For questions or further information, please contact the corresponding author.
