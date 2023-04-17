# Marking-Assistant-V2
22/01/2023

## Summary

This is a JavaScript based project that utilizes the Word JavaScript API to perform action on the current opened document. This project was built to help English teachers grade essays by allowing users to highlight certain sections of the document related to the marking rubric, then compile a report of these highlightings. The users interact with the program using the built in taskpane Add-In functionality provided by the Word API.

## Installation 

Installation currenlty uses sideloading. This method is usually meant for testing of the program but can be used for deployment on a small scale.

- For Ipad, see https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-ipad
- For Windows, see https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins
- For Mac, see https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac
- For Web, see https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing

## Functionality

In the taskpane, there are 10 different buttons that highlight the currently selected section in the open Word document with a specific colour. Each colour corresponds to a different 'comment' for that section, i.e. green represents "good/well done" whereas red represents that that portion of text is a sentence fragment. The 11th button is the "compile report" button. This button activates a sequence where the program iterates through the entire document to find all the portions of text that are highlighted. Once it has found all the highlighted text, it will compile all of the text sections into a report that is inserted at the end of the document. This report is meant to allow the user to get an overview of the comments they left on the essay and remember sections that are well done or need improvements.

## History

This is version 2 of this project. This version is a significant improvement on the last, improving report compiling times by several factors. In its current form, this is the first version of the V2 release. Future updates to come.

## Acknowledgents

This project was built in collaboration with Elaina Mohrmann, an English teacher, who advised on key features and functionality of the program.

## License

MIT License

Copyright (c) 2023 Connor Chan

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
