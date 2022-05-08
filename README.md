# DocXX

#### Introduce
An upgraded version of Duckx that supports adding picture, etc.

#### Status
Create, read and write Microsoft Office Word docx files.
More informations are available in [this](https://docxx.readthedocs.io/en/latest/) documentation.
- Documents (docx) [Word]
	- Read/Write/Edit
	- Change document properties
	- Add picture in Paragraph
	- Add text in Paragraph


#### Install

Easy as pie!

#### Quick Start

Here's an example of how to use duckx to read a docx file; It opens a docx file named **file.docx** and goes over paragraphs and runs to print them:
```c++
#include <iostream>
#include <docxx/duckx.hpp>

int main() {

    docxx::Document doc("file.docx", true);   

    for (auto p : doc.paragraphs())
	for (auto r : p.runs())
            r.add_picture(doc, R"(E:\Temp\Test\view.png)", 110, 140);
}
```
<br/>
And compile your file like this:

```bash
g++ sample1.cpp -ldocxx
```

* See other [Examples](https://github.com/tony2u/DocxX/tree/master/Test)

#### Requirements

- [zip](https://github.com/kuba--/zip)
- [pugixml](https://github.com/zeux/pugixml)


#### Donation


#### Licensing
This library is available to anybody free of charge, under the terms of MIT License (see LICENSE.md).
