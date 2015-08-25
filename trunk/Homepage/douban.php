<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE>something</TITLE>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<style media="all">
	body, html{border:0; margin:0; padding:0; font-size:13px;}
	.entry{width:300px; height:160px; float:left; overflow:hidden;}
	.entry h4{margin:0; padding:0;}
	.entry img{width:100px; height:147px; border:none; float:left; margin:0 10px 10px 0;}
	.entry span{display:block;}
</style>
<style media="print">
	.delete{display:none;}
</style>

</HEAD>
<BODY>

<?php
	require_once("blog/wp-includes/class-snoopy.php");
   	$snoopy = new Snoopy;
	$snoopy->fetch("http://api.douban.com/people/mathzqy/collection?cat=book&status=wish&max-results=100&apikey=7576f60543a81cf56064e20fcaf856df&alt=json");
    // this will copy the created tex-image to your cache-folder
    
    
	$rss_result = str_replace("$", "_", $snoopy->results);
	$rss_result = str_replace("@", "__", $rss_result);
	$rss_result = str_replace("db: ", "", $rss_result);
	$rss_result = str_replace("db:", "", $rss_result);
//	echo $rss_result;
	if (strlen($rss_result) < 100) exit();
		
	$feed = json_decode($rss_result); 
	$entries = $feed->entry;
	
	echo "
		<h2>清单</h2>";
	foreach ($entries as $key=>$entry):
		$subject = $entry->subject;
		$attributes = array_reverse($subject->attribute);
		$attr = array();
		foreach ($attributes as $attribute):
			$attr[$attribute->__name] = $attribute->_t;
		endforeach;
		
		$img = $subject->link[2]->__href;
		$img = str_replace("spic", "mpic", $img);
		$price = $attr['price'];
		$price = str_replace("_", "$", $price);
		
		echo "
		<div class='entry' id='entry$key'>
			<img src='$img' alt='{$subject->title->_t}'/>
			<h4><a href='{$subject->link[1]->__href}'>{$subject->title->_t}</a></h4>
			<p>
				<span>作者: {$attr['author']}</span>
				<span>出版社: {$attr['publisher']}</span>
				<span>出版日期: {$attr['pubdate']}</span>
				<span>价格: $price</span>
				<span class='delete'><a href='javascript:;' onclick='document.getElementById(\"entry$key\").style.display = \"none\";return false;'>取消打印此书信息</a></span>
			</p>
		</div>";
	endforeach;
?>
<script type="text/javascript">
	var feed = {<?php $feed; ?>};
</script>
</BODY>
</HTML>