# PowerPoint Notes Tool

This tool adds notes to your PowerPoint presentation from a text file. If you have a PowerPoint with a ton of slides, you can add your notes to a text file in a certain format, and use this script to add them automatically. 

![PowerPoint Notes Tool](https://github.com/JeremyMorgan/PowerPointNotesTool/blob/main/demo.png)

This is an easy to use script and can save you a lot of time. Here's how to use it. 

## Installation

To install the pre-requisites for this script, you can install

```
pip install python-pptx
```

or 

```
pip install -r requirements.txt
```

## Text File Preparation

Your text file should be delimited with a start and end around the text for each slide. For example:

```
start slide 1
"What you want to say in slide one"
end slide 1

start slide 2
"What you want to say in slide two"
end slide 2
```

Once you've added all your text, you're ready to run the script. 


## Usage

Once you have `Presentation.pptx` in the root folder, and `notes.txt` filled out, run:

```
python app.py
```

It will produce a file, `presentation_with_notes.pptx` that contains all of your notes. 


## Dependencies

This project uses the [python-pptx](https://python-pptx.readthedocs.io/en/latest/) library which features a ton of great features for working with PowerPoint. 

The demo slide deck I used was generated with [beautiful.ai](https://www.beautiful.ai/) which creates amazing presentations based on your input, and can even generate pages with AI (like the demo slide deck). I highly reccomend checking them out. 


