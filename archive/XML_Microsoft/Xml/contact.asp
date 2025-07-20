<%@language="javascript"%>
<%
  var xpath;
  var sName=Request.QueryString("SearchID")();
  if (sName) 
    xpath = "//contact[name='" + sName + "']";
  else 
    xpath = "contacts";

  try {
    var oDs = Server.CreateObject("MSXML2.DOMDocument.3.0");
    oDs.async = false;
    oDs.resolveExternals = false;
    oDs.validateOnParse = false;

    var path = Server.MapPath("contacts.xml"); 
    if ( oDs.load(path) == true ) {
      var oContact= oDs.selectSingleNode(xpath);
      Response.ContentType = "text/xml";
      Response.Write(oContact.xml);
    }
  }
  catch (e) {
    Response.ContentType = "text/xml";
    Response.Write("<error>failed to create Contacts:"
                  +"<desc>"+e.description+"</desc>"
                  +"</error>");
  }
%>