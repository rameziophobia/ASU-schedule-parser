# ASU-schedule-parser
 
user must have a version of python installed  
before usage, please install openpyxl by typing into the terminal  

``` Shell
    $ pip install openpyxl  
```

simple usage  
write the list of course codes in the requested courses list at the top of the script then run the script  
this outputs an excel file of the schedule

``` Shell
    $ py exams.py
```

possible ways of improvement (feel free to contribute)
* accept course codes as command line arguments (ex: use argparse module)
* read from pdf the text content. See [stackoverflow](https://stackoverflow.com/questions/34837707/how-to-extract-text-from-a-pdf-file) 