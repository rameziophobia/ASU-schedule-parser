# ASU-schedule-parser

## installation

user must have a version of python installed  
before usage, please install python package requirements by typing into the terminal  

``` Shell
    $ pip install -r requirements.txt
```

## usage

run the script with space separated list of the requested subject courses as written in the schedule, example:

``` Shell
    $ py exams.py -c "CSE225 CSE325"
```

this outputs an image and an excel file of the schedule  

The file temp1.xlsx acts as a template for the output, you may change its appearance as you like.  
possible ways of improvement (feel free to contribute):

* read from pdf the text content. See [stackoverflow](https://stackoverflow.com/questions/34837707/how-to-extract-text-from-a-pdf-file)
