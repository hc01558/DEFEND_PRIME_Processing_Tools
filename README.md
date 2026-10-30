## DEFEND PRIME Processing Tools
This repository contains the data processing tools used by the DEFEND PRIME study to:
- Perform the required value selection and anonymisation processes on the raw EMIS Web data extract produced by the study's .XML file.
- Convert the raw TPP SystmOne data extract produced by the search and report output configuration available to members of the DEFEND PRIME group in TPP SystmOne, into the aforementioined EMIS Web format.
- Perform the required value selection and anonymisation processes on the TPP SystmOne extract once converted into the EMIS Web format.
- Validate the raw data extracts from EMIS Web and TPP SystmOne.
- Validate the EMIS Web and TPP SystmOne files on which the value selection and anonymisation processes have been performed.

Author: Hugh Claridge<br>
Start date: March 2023<br>
Last edited date: 30th October 2026<br>
Version: v1.0<br>
Version history: Public versions of the DEFEND PRIME study's data processing tools

Reference to the Protocol of the DEFEND PRIME study in BMJ Open:<br>
Claridge _et al_. (2024): Determining the feasibility of calculating pancreatic cancer risk scores for people with new-onset diabetes in primary care (DEFEND PRIME): study protocol<br>
DOI: [10.1136/bmjopen-2023-079863](https://doi.org/10.1136/bmjopen-2023-079863)

## Glossary
MS - Microsoft
VBA - Visual Basic for Applications
XML - eXtensible Markup Language

## Download 
You can clone this project, or use the below link to obtain the files.<br>

## Repository contents

### EMIS Web
| Component | Description | Contents |
|---|---|---|
| [Deployed workbook](https://github.com/hc01558/DEFEND_PRIME_Processing_Tools/tree/main/DEFEND_PRIME_Processing_Tools/EMIS%20Web/Deployed_Workbook) | Ready to use EMIS Web extract processing tool with instructions. | Excel workbook with VBA modules installed; detailed Word user instructions. |
| [Source VBA modules](https://github.com/hc01558/DEFEND_PRIME_Processing_Tools/tree/main/DEFEND_PRIME_Processing_Tools/EMIS%20Web/Source_VBA_Modules) | Source code for the EMIS Web extract processing software, provided separately. | Exported VBA `.bas` modules as installed in the deployed workbook. |
| [Extract validation](https://github.com/hc01558/DEFEND_PRIME_Processing_Tools/tree/main/DEFEND_PRIME_Processing_Tools/EMIS%20Web/Validation_Workbooks/Extract_Validation) | Validates the structure and content of raw EMIS Web extracts for extraction development. | Excel validation workbook containing a test worksheet and synthetic data to demonstrate the validation process. |
| [Processing validation](https://github.com/hc01558/DEFEND_PRIME_Processing_Tools/tree/main/DEFEND_PRIME_Processing_Tools/EMIS%20Web/Validation_Workbooks/Processing_Validation) | Validates the EMIS Web extract processing software, checking that processing produces the expected results. | Excel validation workbook containing a test worksheet and synthetic data to demonstrate the validation process; Word documentation describing the tests; separate Word documentation of virtual machine testing procedures. |
 
### TPP SystmOne
| Component | Description | Contents |
|---|---|
| [Deployed workbook](https://github.com/hc01558/DEFEND_PRIME_Processing_Tools/tree/main/DEFEND_PRIME_Processing_Tools/TPP%20SystmOne/Deployed_Workbook) | Ready to use TPP SystmOne extract processing tool with instructions. | Excel workbook with VBA modules installed; detailed Word user instructions. |
| [Source VBA Modules](https://github.com/hc01558/DEFEND_PRIME_Processing_Tools/tree/main/DEFEND_PRIME_Processing_Tools/TPP%20SystmOne/Source_VBA_Modules) | Source code for the TPP SystmOne extract processing software, provided separately. | Exported VBA `.bas` modules as installed in the deployed workbook. |
| [Extract validation](https://github.com/hc01558/DEFEND_PRIME_Processing_Tools/tree/main/DEFEND_PRIME_Processing_Tools/TPP%20SystmOne/Validation_Workbooks/Extract_Validation) | Validates the structure and content of raw TPP SystmOne extracts for extraction development. | Excel validation workbook containing a test worksheet and synthetic data to demonstrate the validation process. |
| [Processing validation](https://github.com/hc01558/DEFEND_PRIME_Processing_Tools/tree/main/DEFEND_PRIME_Processing_Tools/TPP%20SystmOne/Validation_Workbooks/Processing_Validation) | Validates the TPP SystmOne to EMIS Web harmonisation software and processing software, checking that processing produces the expected results. | Excel validation workbook containing two test worksheets and synthetic data to demonstrate the validation process; Word documentation describing the tests; separate Word documentation of virtual machine testing procedures. |
