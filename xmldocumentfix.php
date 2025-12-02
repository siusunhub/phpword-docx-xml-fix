<?php

$process_log = array();
$show_log = false;

function savelog($log)
{
	$GLOBALS["process_log"][] = $log;
	if ($GLOBALS["show_log"]) echo $log . "<br>";
}


function removetmp()
{

	if (file_exists($GLOBALS["tmpfolder"])) {
		if (file_exists($GLOBALS["tmpfolder"] . "document.xml"))
			@unlink($GLOBALS["tmpfolder"] . "document.xml");
		if (file_exists($GLOBALS["tmpfolder"] . "document.fix.xml"))
			@unlink($GLOBALS["tmpfolder"] . "document.fix.xml");
		if (file_exists($GLOBALS["tmpfolder"] . $GLOBALS["new_docx_tmp_file"]))
			@unlink($GLOBALS["tmpfolder"] . $GLOBALS["new_docx_tmp_file"]);

		@rmdir($GLOBALS["tmpfolder"]);
	}
}


function output_download_header($_new_filename, $mime_type = "", $filesize = "")
{

	if (preg_match("/Android/", $_SERVER['HTTP_USER_AGENT'])) {
		header("Content-Type: application/octet-stream");
	} else {
		if ($mime_type == "") $mime_type = 'application/octetstream';
		header('Content-Type: ' . $mime_type);
	}

	$encoded_filename = urlencode($_new_filename);
	$encoded_filename = str_replace("+", "%20", $encoded_filename);

	if (preg_match("/MSIE/", $_SERVER['HTTP_USER_AGENT'])) {
		header('Content-Disposition:  attachment; filename="' . $encoded_filename . '"');
	} elseif (preg_match("/Firefox/", $_SERVER['HTTP_USER_AGENT'])) {
		header('Content-Disposition: attachment; filename*="utf8' .  $_new_filename . '"');
	} elseif (preg_match("/Android/", $_SERVER['HTTP_USER_AGENT'])) {
		$info = pathinfo($encoded_filename);
		$_new_filename =  basename($encoded_filename, "." . $info['extension']) . '.' . strtoupper($info['extension']);

		//						$_new_filename = "abc.TIF";
		//						fwrite($fp2, "\n" . $_new_filename ."\n");

		header('Content-Disposition: attachment; filename="' .  $_new_filename . '"');
	} else {
		header('Content-Disposition: attachment; filename="' .  $_new_filename . '"');
	}
	//					fwrite($fp2, "\nB\n");

	if ($filesize > 0)   header('Content-Length: ' . $filesize);
	header('Expires: 0');
	header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
	header('Pragma: public'); //   no-cache


}

function output_download_file($_url, $filename, $mine_type, $exit = true)
{
	if (file_exists($_url)) {

		$_file_size = filesize($_url);
		output_download_header($filename, $mime_type, $_file_size);
		readfile($_url);

		if ($exit) exit;
	}
	return false;
}


if ($_POST["submit"]) {

	savelog("Processing file " . $_FILES["file"]["tmp_name"]);

	$_source_docx_file = $_FILES["docxfile"]["tmp_name"];
	$_source_docx_filename = $_FILES["docxfile"]["name"];
	$tmpfolder = "/tmp/docx_" . rand(10000, 99999) . "/";
	$new_docx_tmp_file = "newdocx_" . rand(10000, 99999) . ".docx";

	savelog("Opening file " . $_source_docx_filename);

	$zip = new ZipArchive;
	$res = $zip->open($_source_docx_file);
	if ($res === TRUE) {
		savelog("Create folder " . $tmpfolder);
		mkdir($tmpfolder);
		savelog("Extract document.xml");

		$zip->extractTo($tmpfolder, array('word/document.xml'));
		$zip->close();

		if (file_exists($tmpfolder . 'word/document.xml')) {

			savelog("Move document.xml");
			rename($tmpfolder . 'word/document.xml', $tmpfolder . 'document.xml');
			rmdir($tmpfolder . 'word/');

			savelog("Opening document.xml");

			$output = "";
			$handle = fopen($tmpfolder . "document.xml", "r");

			$fix_ct = 0;
			$buffer = "";

			savelog("Reading files");

			while (!feof($handle)) {
				$buffer .= fread($handle, 8192);
				$loop = true;
				$have_end = true;
				$lastpos = 0;
				if ($buffer != "") {
					do {
						if ($lastpos < strlen($buffer)) {

							$st = strpos($buffer, '$' . '<', $lastpos);
							if ($st !== false) {
								$en = strpos($buffer, '{', $st);
								if ($en !== false) {
									$substring = substr($buffer, $st, $en - $st + 1);

									$nontag = strip_tags($substring);
									if ($nontag != $substring) $fix_ct++;
									$buffer = substr($buffer, 0, $st) . $nontag . substr($buffer, $en + 1);

									if ($nontag != $substring) {
										savelog("Convert1 :" . htmlspecialchars($substring) . " to " . htmlspecialchars($nontag));
									} else {
										if ($_POST["showdebugall"])
											savelog("Located1 :" . htmlspecialchars($nontag));
									}

									$lastpos = $st + 1;
								} else {
									$have_end = false;
									$loop = false;
								}
							} else {
								$loop = false;
								$have_end = true;
							}
						} else {
							savelog("error offset1 : lastpos:" . $lastpos . "/len:" . strlen($buffer));
							$loop = false;
							$have_end = true;
						}
					} while ($loop);

					$loop = true;
					$lastpos = 0;

					do {
						if ($lastpos < strlen($buffer)) {
							$st = strpos($buffer, '$' . '{', $lastpos);
							if ($st !== false) {
								$en = strpos($buffer, '}', $st);
								if ($en !== false) {
									$substring = substr($buffer, $st, $en - $st + 1);
									$nontag = strip_tags($substring);
									if ($nontag != $substring) $fix_ct++;
									$buffer = substr($buffer, 0, $st) . $nontag . substr($buffer, $en + 1);
									$lastpos = $st + 1;

									if ($nontag != $substring) {
										savelog("Convert2 :" . htmlspecialchars($substring) . " to " . htmlspecialchars($nontag));
									} else {
										if ($_POST["showdebugall"])
											savelog("Located2 :" . htmlspecialchars($nontag));
									}
								} else {
									$have_end = false;
									$loop = false;
								}
							} else {
								$loop = false;
								$have_end = true;
							}
						} else {
							savelog("error offset2 : lastpos:" . $lastpos . "/len:" . strlen($buffer));
							$loop = false;
							$have_end = true;
						}
					} while ($loop);
					if ($have_end) {
						$output .= $buffer;
						$buffer = "";
					}
				}
			}

			fclose($handle);
			$output .= $buffer;

			savelog("Writing");

			// write buffer;
			$handle2 = fopen($tmpfolder . "document.fix.xml", "w");
			fwrite($handle2, $output);
			fclose($handle2);

			savelog("Total fix: " . $fix_ct);

			if ($fix_ct > 0) {
				savelog("Copy docx to new file: " . $fix_ct);
				copy($_source_docx_file, $tmpfolder . $new_docx_tmp_file);

				savelog("Open new temp docx file");

				$zip = new ZipArchive;
				$res = $zip->open($tmpfolder . $new_docx_tmp_file);
				if ($res === TRUE) {

					savelog("Add document.xml");
					$zip->addFile($tmpfolder . "document.fix.xml", 'word/document.xml');
					$zip->close();

					if ($_POST["showdebug"] == "") {
						output_download_file($tmpfolder . $new_docx_tmp_file, $_source_docx_filename, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", false);
					}

					savelog("complete. remove temp");
					removetmp();

					if ($_POST["showdebug"] == "") {
						exit;
					}
				} else {
					savelog("Error open zip");
					removetmp();
				}
			} else {
				savelog("No change, remove temp file");
				removetmp();
			}
		} else {
			savelog("Error extract xml file ");
			removetmp();
		}
	} else {
		savelog("Error open zip");
		removetmp();
	}
}


?>
<html>

<body>
	<form method="post" enctype="multipart/form-data" name="form1" id="form1">
		<strong>DOCX XML Fix</strong><br>
		<label for="fileField">Select docx file:</label>
		<input type="file" name="docxfile" id="docxfile"><br>
		<input type="checkbox" id="showdebug" name="showdebug" value="1">
		<label for="showdebug">Show Debug Only</label><br>
		<input type="checkbox" id="showdebugall" name="showdebugall" value="1">
		<label for="showdebugall">Show All Template Tag</label><br>
		<input type="submit" name="submit" id="submit" value="Scan and Fix">
		<input type="submit" name="reset" id="reset" value="Clear">
	</form>
	<?php
	if (is_array($process_log) && count($process_log) > 0) {
		echo "<hr><pre>";
		foreach ($process_log as $_l) {
			echo  $_l . "\n";
		}
		echo "</pre>";
	}
	?>
</body>

</html>