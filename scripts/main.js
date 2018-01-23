var myHeading = document.querySelector('h1');
myHeading.textContent = 'Hello world!';

var connection = new ActiveXObject("ADODB.Connection");

var connectionstring="Data Source=liuxiao-PC;Initial Catalog=marketplace_lending_staging_tables;Integrated Security=SSPI;Provider=SQLOLEDB";

connection.open(connectionstring);

var rs = new ActiveXObject("ADODB.Recordset");

str_query = "SELECT TOP 10 [借款人编号],[借款编号] FROM [marketplace_lending_staging_tables].[dbo].[zhengda.oct2017.loantape]"

rs.Open(str_query, connection);
rs.MoveFirst();
while(rs.EOF != true)
{
   document.write(rs.fields(1));
   alert(rs("借款人编号") + "\t" + rs("借款编号"));  
   rs.MoveNext();
}


/*rs.MoveFirst();
var body = document.getElementsByTagName('body')[0];
var tbl = document.createElement('table');
tbl.style.width = '100%';
tbl.setAttribute('border', '1');
var tbdy = document.createElement('tbody');
for (var i = 0; i < 3; i++) {
    var tr = document.createElement('tr');
    for (var j = 0; j < 2; j++) {
        if (i == 2 && j == 1) {
            break
        } else {
            var td = document.createElement('td');
            td.appendChild(document.createTextNode("Cell " + i + "," + j + rs.fields(1)))
            rs.MoveNext();
            i == 1 && j == 1 ? td.setAttribute('rowSpan', '2') : null;
            tr.appendChild(td)
        }
    }
    tbdy.appendChild(tr);
}
tbl.appendChild(tbdy);
body.appendChild(tbl)*/

rs.close();
connection.close();
