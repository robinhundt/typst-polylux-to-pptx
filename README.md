# Typst Polylux to PPTX

This is a small script to convert a typst presentation which uses polylux to a powerpoint .pptx presentation. The typst presentation is compiled to .png and then python-pptx is used to create a presentation where each image becomes one slide. The speaker notes are also extracted from the typst presentation and added to the powerpoint.

To use this, install the dependencies with poetry (or your package manager of choice) and change the `typs_presentation` path in `main.py`. You also need a local installation of the typst compiler and the `polylux2pdfpc` tool (available via AUR or the source from [here](https://github.com/andreasKroepelin/polylux/tree/main/pdfpc-extractor)).
