## Welcome to GitHub Pages

You can use the [editor on GitHub](https://github.com/liutrees/codedemo/edit/master/index.md) to maintain and preview the content for your website in Markdown files.

Whenever you commit to this repository, GitHub Pages will run [Jekyll](https://jekyllrb.com/) to rebuild the pages in your site, from the content in your Markdown files.

### Markdown

Markdown is a lightweight and easy-to-use syntax for styling your writing. It includes conventions for

```markdown
Syntax highlighted code block
<script type="text/javascript">
        function ReadExcel() {
           var tempStr = "";
           var temps="";
           var reg="";
           //得到文件路径的值
            var mainpath=document.URL;
           // var path=document.all.excelpath.value+"\\"+"test.xlsx";
        //alert(document.all.excelpath.value);
        var num=document.getElementById("sheet").value;
        //alert(path);
         var stunum=document.getElementById("stunum").value;
          //alert(stunum);
           //创建操作EXCEL应用程序的实例
           var oXL = new ActiveXObject("Excel.application");
            //打开指定路径的excel文件
           //var oWB = oXL.Workbooks.open("C:\\Users\\Administrator\\Desktop\\获取excel的行和列\\获取excel的行和列\\test.xlsx");
         var oWB = oXL.Workbooks.open("http://liutrees.3vkj.net/test.xlsx");
         //  var oWB = oXL.Workbooks.open("\\test.xlsx");
 
          //操作第一个sheet(从一开始，而非零)
           oWB.worksheets(parseInt(num)).select();
           var oSheet = oWB.ActiveSheet;
           //使用的行数
         var rows =  oSheet .usedrange.rows.count; 
           try {
              for (var i = 2; i <= rows; i++) {
           //  if (oSheet.Cells(i, 2).value == "null" || oSheet.Cells(i, 3).value == "null") break;
           //     var a = oSheet.Cells(i, 2).value.toString() == "undefined" ? "": oSheet.Cells(i, 2).value;
               if(oSheet.Cells(i, 1).value.localeCompare(stunum)==0){
               tempStr += (" " + oSheet.Cells(i, 1).value + " " + oSheet.Cells(i, 2).value + " " + oSheet.Cells(i, 3).value + " " + oSheet.Cells(i, 4).value + "\n"); 
	break;
                      }
              }
               if(i>rows)
                    tempStr="查无信息，班级有没有选对！";
           } catch(e) {
              document.getElementById("txtArea").value = tempStr;
           }
           document.getElementById("txtArea").value = tempStr;
           //退出操作excel的实例对象
           oXL.Application.Quit();
            //手动调用垃圾收集器
           CollectGarbage();
        }
  </script>

学号：
<input type="text"  id="stunum" value=""></textarea>
<select id="sheet">
  <option value ="1">英语1801</option>
  <option value ="2">英语1802</option>
  <option value ="3">传播1803</option>
  <option value ="4">汉语1801</option>
  <option value ="5">历史1801</option>
  <option value ="6">其他</option>
</select>
<input type="button" onclick="ReadExcel();" value="查询">
<br>
<textarea id="txtArea" cols=50 rows=10></textarea>
# sdfsd
## Header 2
### Header 3

- Bulleted
- List

1. Numbered
2. List

**Bold** and _Italic_ and `Code` text

[Link](url) and ![Image](src)
```

For more details see [GitHub Flavored Markdown](https://guides.github.com/features/mastering-markdown/).

### Jekyll Themes

Your Pages site will use the layout and styles from the Jekyll theme you have selected in your [repository settings](https://github.com/liutrees/codedemo/settings). The name of this theme is saved in the Jekyll `_config.yml` configuration file.

### Support or Contact

Having trouble with Pages? Check out our [documentation](https://help.github.com/categories/github-pages-basics/) or [contact support](https://github.com/contact) and we’ll help you sort it out.
