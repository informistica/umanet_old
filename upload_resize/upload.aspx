<%@ Page Language="C#" AutoEventWireup="True" Debug="true" %>
<%@ Import Namespace="System" %> 
<%@ Import Namespace="System.IO" %> 
<%@ Import Namespace="System.Drawing" %> 
<%@ Import Namespace="System.Drawing.Imaging" %> 
<%@ Import Namespace="System.Drawing.Drawing2D" %> 
<%@ Import Namespace="System.Collections" %> 
<%@ Import Namespace="System.Runtime.InteropServices" %> 
<%@ Import Namespace="System.Globalization" %> 
<%@ Import Namespace="System.Web.UI.HtmlControls" %>

<script Language="C#" runat="server" src="wbresize.cs"></script>
<script Language="C#" runat="server" src="quantizer.cs"></script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
	<title>Upload</title>
	<meta name="author" content="Danilo Cicognani" />
	<meta name="robots" content="noindex,nofollow"/>
    
    <!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
     <link rel="stylesheet" href="../../css/style-themes.css">
	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script> 
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
    
	<script runat="server">
		string UploadName, FileName, ThumbFile;
		bool Upload = false;

		void UploadBtn_Click(Object sender, EventArgs e) {
			// Display information about posted file
			UploadName = fileupload.PostedFile.FileName;
			int i = UploadName.LastIndexOf("\\");
			FileName = "";
			if (i != 0)
				FileName = UploadName.Substring(i + 1);
			else
				FileName = UploadName;

			int posDot = FileName.LastIndexOf(".");
			string FileNameWithoutExtension = FileName.Substring(0, posDot);
			string Extension = "jpg";
			switch (fileupload.PostedFile.ContentType) {
				case "image/gif":
					Extension = ".gif";
					break;
				case "image/pjpeg":
					Extension = ".jpg";
					break;
				case "image/x-png":
					Extension = ".png";
					break;
			}

			bool toResize = false;
			string FileNameWithoutExtensionJpg = FileNameWithoutExtension;
			if (imgWidth.Value != "")
				toResize = true;
			if (FileName != prev.Value) {
				int j = 1;
				while ( File.Exists(path.Value + "\\" + FileName) || (toResize && File.Exists(path.Value + "\\" + FileNameWithoutExtensionJpg + ".jpg")) ) {
					FileName = FileNameWithoutExtension + j + Extension;
					FileNameWithoutExtensionJpg = FileNameWithoutExtension + j;
					j = j + 1;
				}
				if (j > 1)
					FileNameWithoutExtension = FileNameWithoutExtension + (j - 1);
			}

			fileupload.PostedFile.SaveAs(path.Value + "\\" + FileName);
			if (thumbWidth.Value != "") {
				//I must create the thumbnail
				double requestedWidth = Convert.ToDouble(thumbWidth.Value);
				double requestedHeight = (thumbHeight.Value == null || thumbHeight.Value == "") ? 0 : Convert.ToDouble(thumbHeight.Value);
				ResizeImage(path.Value + "\\" + FileName, thumbPath.Value + "\\" + FileNameWithoutExtension + ".jpg", "jpg", requestedWidth, requestedHeight);
				ThumbFile = FileNameWithoutExtension + ".jpg";
			}
			if (imgWidth.Value != null && imgWidth.Value != "") {
				//I must resize the original image
				double requestedWidth = Convert.ToDouble(imgWidth.Value);
				double requestedHeight = (imgHeight.Value == null || imgWidth.Value == "") ? 0 : Convert.ToDouble(imgWidth.Value);
				ResizeImage(path.Value + "\\" + FileName, path.Value + "\\" + FileNameWithoutExtension + ".jpg", "jpg", requestedWidth, requestedHeight);
				if (Extension != ".jpg") {
					//Delete the original file
					//File.Delete(path.Value + "\\" + FileName);
				}
				FileName = FileNameWithoutExtension + ".jpg";
			}
			// Save uploaded file to server
		    Session["FileName"]=FileName;
			//creo file di testo con il nome del file dentro 
			 StreamWriter fp;
			  try
			  {
				//DeleteFile(Server.MapPath("FileName.txt"))
				fp = File.CreateText(Server.MapPath("FileName.txt"));
				fp.WriteLine(FileName);
				fp.Close();
			  }
			  catch (Exception err)
			  {
				
			  }
						
			
			
			Upload = true;
		}

		string jsEncode(string js) {
			string toRet = "";
			if (js != null) {
				toRet = js.Replace("\\", "\\\\");
				toRet = toRet.Replace("'", "\\'");
			}
			return toRet;
		}

		//'Resizes the image FileName, to file OutFileName, to OutFormat, to Width and Height specified, preserving aspect ratio
		void ResizeImage(string pathFile, string OutFileName, string OutFormat, double Width, double Height) {
			try {
				wbResize ImageResizer = new wbResize();

				ImageResizer.LoadImage(pathFile);
				ImageResizer.SetFileType(OutFormat);
				ImageResizer.SetThumbSize(Width, Height, false);

				ImageResizer.SetOptionHighQuality(true);
				ImageResizer.SetOptionJpegQuality(100);
				ImageResizer.SetOptionFilter(InterpolationMode.HighQualityBicubic);
				ImageResizer.SetOptionGifDepth(4);
				ImageResizer.SetOptionGifPalette(255);
				ImageResizer.SetOptionTiffCompress(true);

				ImageResizer.ResizeAndSave(OutFileName);
			}
			catch {
				Response.Write("<span class=\"error\">" + OutFileName + " resize error</span><br/>");
				Response.End();
			}
		}
	</script>

<% 
if (Upload == true) {
	frmUpload.Visible = false;
%>
	<script language="javascript">
		window.opener.document.<%= field.Value %>.value = '<%= jsEncode(FileName) %>';
		<% if (thumbWidth.Value != "") { %>
		window.opener.document.<%= thumbField.Value %>.value = '<%= jsEncode(ThumbFile) %>';
		<% } %>
		window.close();
	</script>
<%
}
else {
	field.Value = Request.QueryString["field"];
	path.Value = Request.QueryString["path"];
	prev.Value = Request.QueryString["prev"];
	thumbField.Value = Request.QueryString["thumbField"];
	thumbPath.Value = Request.QueryString["thumbPath"];
	thumbWidth.Value = Request.QueryString["thumbWidth"];
	thumbHeight.Value = Request.QueryString["thumbHeight"];
	imgWidth.Value = Request.QueryString["imgWidth"];
	imgHeight.Value = Request.QueryString["imgHeight"];
}
%>
  <script language="javascript" type="text/javascript" >
 function validate2() {
	var stringa=frmUpload.fileupload.value;
	//alert(stringa); 
	if ((stringa.search(".jpg") == -1) && (stringa.search(".JPG") == -1))
	{
	   alert("L'immagine deve essere in formato .jpg");
	   frmUpload.fileupload.setfocus();
	   return 0;
	}
 	else
	{
	    document.frmUpload.action = "upload.asp";
		document.frmUpload.submit();		
   }
}
</script>

</head>
<body>
	<h3>Scegli foto (solo .jpg)</h3>
	<form id="frmUpload" action="upload.aspx" method="post" enctype="multipart/form-data" runat="server">
		<input type="hidden" id="field" value="" runat="server"/>
		<input type="hidden" id="path" value="" runat="server"/>
		<input type="hidden" id="prev" value="" runat="server"/>
		<input type="hidden" id="thumbField" value="" runat="server"/>
		<input type="hidden" id="thumbPath" value="" runat="server"/>
		<input type="hidden" id="thumbWidth" value="" runat="server"/>
		<input type="hidden" id="thumbHeight" value="" runat="server"/>
		<input type="hidden" id="imgWidth" value="" runat="server"/>
		<input type="hidden" id="imgHeight" value="" runat="server"/>
		<input id="fileupload" type="file" runat="server" class="btn-primary"/>
		<input id="submit" type="submit" value="Invia" onserverclick="UploadBtn_Click" class="btn" runat="server"/>
	</form>
</body>

 <script type="text/javascript">
	 
$(window).load(function () {
	   
	  $('#fileupload').click();
	  
	 
	    event.stopPropagation();
	    
	});
	
</script>

</html>