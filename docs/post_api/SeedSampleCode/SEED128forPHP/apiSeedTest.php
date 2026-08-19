<?
require_once ('SEED128.php');

function fixNull($source){
      $rData = "";
      if ($source == NULL) {
      	$rData = "";
      }else{ 
      	$rData = $source;
      }
      return $rData;	
}

		$seedKey = fixNull($_POST['skey']);
		$data = "";
		$encryptStr = "";
		
		$data = fixNull($_POST['data']);
		$encryptStr = fixNull($_POST['encryptStr']);
		
		$seed = new SEED128();
		if ($data!=NULL) {
			$encryptStr = $seed->getEncryptData($seedKey, $data);
		} else if ($encryptStr!=NULL){
			$data = $seed->getDecryptData($seedKey, $encryptStr);
		}

?>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"> 
<title>SEED 테스트</title>
</head>	
<style>
TD {font-family:"Dotum";font-size:9pt;color:#535353}
INPUT {vertical-align:middle;font-size:9pt;font-family:"Dotum";border-color:#99aadd; border-style:solid;border-width: 1px; height:19px; background-color:#f3f3f3}
</style>
<body>

<form method="post" name="frm">

<table cellpadding="4" width="800">
<tr><td align="center"><input type="submit"></td></tr>
</table>


<h1>Query URL</h1>
seedKey : <input type="text" size="30" name="skey" value="<?=$seedKey?>"><br><br>
평문<br>
<textarea name="data" cols="80" rows="17" ><?=$data?></textarea>
<br>
<br>
암호문<br>
<textarea name="encryptStr" cols="80" rows="17"><?=$encryptStr?></textarea>
</form>
</body>
</html>

