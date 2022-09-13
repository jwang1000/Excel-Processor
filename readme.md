# Excel Processor

Current version: v0.0.1

## About and Usage

An easy to use application that takes data from an Excel sheet and inserts it into a Word document.

Each group of cells will replace a tag that is specified in curly braces, i.e. {replace me}. 

Multiple of the same tag can be used in the Word document, and the same data will be copied for each duplicate tag. Any tags that are in the Word document but not given any data from the Excel sheet will be replaced with `undefined`.

You can specify which cells to select, the name of each tag, the name of the output file, and save and load presets to easily repeat your work.


## Installation

Download the installer. Unzip the folder and run the `Excel Processor-0.0.1 Setup.exe` file.

The app will open automatically once it is finished installing, and there will also be a shortcut installed on the desktop.

Under the terms of the Apache 2.0 license: commercial use, modification, distribution, patent use, and private use are all permitted. There is no liability against the creators/developers of the software and no warranty provided. [See the terms of the License](./LICENSE) for the full details.


## Building Manually

To make changes or build the code manually, clone the repository. 
Make sure you have [Node.js](https://nodejs.org/) installed. 
Then inside the folder, open a terminal and run `npm install` and then `npm start`.

To create an installer, run `npm run make`.

Changes are welcome via pull request! Excel Processor uses the Apache 2.0 open source license.