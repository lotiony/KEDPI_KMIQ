<%@ Page Language="C#" Inherits="System.Web.Mvc.ViewPage<dynamic>" %>

<!DOCTYPE html>

<html>
<head runat="server">
    <meta name="viewport" content="width=device-width" />
    <title>ResultTest</title>
</head>
<body>
    <form id="form1" name="form1" method="post" action="api/getresult" enctype="application/x-www-form-urlencoded">

        ID : <input type="Text" name="ID" value="test" /><br />
        TypeId : <input type="Text" name="TypeId" value="4" /><br />
        Token : <input type="Text" name="Token" size="35" value="BEAADECD-AD35-4BC1-8962-1A0DC3DAA615<%//Guid.NewGuid() %>" /><br />
        ResultStr : <input type="Text" name="ResultStr" size="150" value="1,2,5,5,5,1,1,1,1,1,2,2,2,2,2,4,3,2,1,1,1,2,3,4,5,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,1,2,3,4,5" /><br />
        Name :  <input type="Text" name="uName" value="테스트" /><br />
        Birth :  <input type="Text" name="uBirth" value="" /><br />
        Email :  <input type="Text" name="uEmail" value="dev@smartomr.com" /><br />
        Tel :  <input type="Text" name="uTel" value="010-3959-1031" /><br />

        <input type="submit" value="Submit" />
    </form>

</body>
</html>
