# docx_deformer
Detect and remove form widgets from docx files

For text mining purposes it can be useful to extract text from .doc and .docx files, which was a procedure that we used in a particular use case (analysing court judgements[1]) . However, we found that the use of templated elements such as pulldown menus made this extremely problematic to achieve, since some commonplace tools for extracting text do not know about these. Therefore we came up with a two-step fix:

step 1: doc-docx convert in libreoffice
step 2: remove multiple-choice pulldowns/calculated fields from docx using calcfields-resolver.py

This is not the last word on the subject, obviously, but it's available in case anyone else discovers a similar problem. 

[1] Scheinert Idodo, L., & Tonkin, E. L. (Accepted/In press). Towards Ethical Judicial Analytics: Assessing Readability of Immigration and Asylum Decisions in the United Kingdom. In Proceedings of the 20th EPIA Conference on Artificial Intelligence Springer.
