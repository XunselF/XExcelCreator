# XExcelCreator

快速创建Excel表格工具


### 初始化

```
	XExcelCreator creator = new XExcelCreator(filename,sheetname,0);
```

### 快速建表

```
	XExcelCreator creator = new XExcelCreator(filename,sheetname,0)	 
		.createText("HelloWorld")								
		.writeData();
```

### 创建多个数据

```
	creator.createAllTexts(new String[]{"1","2","3"});
	creator.writeData();

```

### 其他属性

```
	creator.setmHeight(30);
	creator.setmWidth(300);

	creator.setFontStyle(fontStyle,fontSize,ifFontBlod);
    
```

### License

```

MIT License

Copyright (c) 2018 Shuxun Guo

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

```