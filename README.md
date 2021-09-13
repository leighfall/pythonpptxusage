# pythonpptxusage.py

Program generates a couple of slides, showing usage of the python-pptx library.

## Requirements

- Latest version of Python.
- The python-pptx library. If you do not have this libary, you can install it by completing the following command:

```bash
pip install python-pptx
```

Note: This program generates a pptx file, but the file can be opened online using (at minimum) PowerPoint or Google Slides. It does not automatically open the file for you. Therefore, a desktop version of PowerPoint is not required.

The PowerPoint will need to already exist, and the theme you want installed should already be installed. A created PowerPoint with any "name" that has one blank slide with desired theme will do. The program will prompt for the file's name [name.pptx] and will add all the necessary slides to that file with the theme picked. You will need to delete the original blank slide as the program adds slides to an existing presentation, not replaces.

Ensure that the program and images are in the same folder as the powerpoint created.

### Supporting Libraries
- datetime

## Usage

```bash
python3 main.py
```

```bash
# usage: [name_of_powerpoint.pptx]
# Enter the name of a PowerPoint presentation to be generated: 
```

The file you enter will be updated and saved.

## Issues
- Need a workaround for border creation with tables
- When first slide is created, cannot be accessed using prs.slides[0] as per documentation: https://python-pptx.readthedocs.io/en/latest/api/presentation.html
