<html>
<head>
<title>Upload File with ColdFusion</title>
</head>
<cfif isdefined("form.upload_now")>
<cffile action="upload" filefield="ul_path" destination="d:\hosting\cgvaorg\CGVADOCS" accept="image/jpeg, image/gif" nameconflict="makeunique">
The file was successfully uploaded!
</cfif>

<form action="gd_upload.cfm" method="post" name="upload_form" enctype="multipart/form-data" id="upload_form">
<input type="file" name="ul_path" id="ul_path">
<input type="submit" name="upload_now" value="submit">
</form>
</body>
</html>

