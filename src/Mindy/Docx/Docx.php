<?php

namespace Mindy\Docx;

use Exception;
use Mindy\Utils\RenderTrait;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use ZipArchive;

class Docx
{
    use RenderTrait;

    private $content;
    /**
     * @var string temporary directory
     */
    private $tmpDir;
    private $images = [];

    public function render($docxTemplateFile, array $data = [], array $images = [])
    {
        $this->tmpDir = $this->makeTmpDir();
        if (file_exists($docxTemplateFile) === false) {
            throw new Exception("File not found");
        }

        $this->extract($docxTemplateFile, $this->tmpDir);

        $path = realpath($this->tmpDir . '/word/document.xml');
        $tmpContent = file_get_contents($path);
        // TODO This fix really needed?
        $tmpFixedContent = str_replace("&", "&amp;", $tmpContent);
        $this->content = self::renderString($tmpFixedContent, $data);

        foreach ($images as $ref => $image) {
            $this->addImage($ref, $image);
        }

        return $this;
    }

    /* the ref will be used to assign this image later
     * You can use whatever you want
     * */
    private function addImage($ref, $file)
    {
        if (file_exists($file)) {
            if (!array_key_exists($ref, $this->images)) {
                $this->images[$ref] = $file;
            } else {
                throw new Exception("the ref $ref allready exists");
            }
        } else {
            throw new Exception($file . ' does not exist');
        }
    }

    private function makeTmpDir()
    {
        $path = tempnam(sys_get_temp_dir(), 'docxgen_');
        unlink($path);
        mkdir($path);
        return $path;
    }

    /**
     * @param $outputFile
     * @throws Exception
     * @return bool
     */
    public function save($file)
    {
        $this->processImages();
        file_put_contents($this->tmpDir . '/word/document.xml', $this->content);
        return $this->compact($file);
    }

    private function processImages()
    {
        if (count($this->images) > 0) {
            if (!is_dir($this->tmpDir . "/word/media")) {
                mkdir($this->tmpDir . "/word/media");
            }

            $relationships = file_get_contents($this->tmpDir . "/word/_rels/document.xml.rels");

            foreach ($this->images as $ref => $file) {
                $xml = '<Relationship Id="phpdocx_' . $ref . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' . basename($file) . '" />';
                copy($file, $this->tmpDir . "/word/media/" . basename($file));
                $relationships = str_replace('ships">', 'ships">' . $xml, $relationships);
            }
            file_put_contents($this->tmpDir . "/word/_rels/document.xml.rels", $relationships);
        }
    }

    /**
     * @return bool
     */
    public function extract($docxFile, $dir)
    {
        if (file_exists($dir) && is_dir($dir)) {
            //clean up of the tmp dir
            $this->rrmdir($dir);
            mkdir($dir);
        }

        $zip = new ZipArchive();
        if ($zip->open($docxFile, ZipArchive::CHECKCONS)) {
            $zip->extractTo($dir);
            $zip->close();
            return true;
        } else {
            return false;
        }
    }

    private function compact($output)
    {
        $zip = new ZipArchive();
        if ($zip->open($output, ZipArchive::CREATE | ZipArchive::OVERWRITE)) {

            // Create recursive directory iterator
            /** @var \SplFileInfo[] $files */
            $files = new RecursiveIteratorIterator(
                new RecursiveDirectoryIterator($this->tmpDir),
                RecursiveIteratorIterator::LEAVES_ONLY
            );

            foreach ($files as $name => $file) {
                // Skip directories (they would be added automatically)
                if (!$file->isDir()) {
                    // Get real and relative path for current file
                    $filePath = $file->getRealPath();
                    $relativePath = substr($filePath, strlen($this->tmpDir) + 1);
                    // Add current file to archive
                    $zip->addFile($filePath, $relativePath);
                }
            }

            return $zip->close();
        } else {
            throw new Exception("Failed to create archive");
        }
    }

    private function rrmdir($dir)
    {
        if (is_dir($dir)) {
            $objects = scandir($dir);
            foreach ($objects as $object) {
                if ($object != "." && $object != "..") {
                    if (filetype($dir . "/" . $object) == "dir") $this->rrmdir($dir . "/" . $object); else unlink($dir . "/" . $object);
                }
            }
            reset($objects);
            rmdir($dir);
        }
    }

    private function clean($content)
    {
        $content = str_replace('<w:lastRenderedPageBreak/>', '', $content); // faster
        $content = $this->cleanTag($content, ['<w:proofErr', '<w:noProof', '<w:lang', '<w:lastRenderedPageBreak']);
        $content = $this->cleanRsID($content);
        $content = $this->cleanDuplicatedLayout($content);
        return $content;
    }

    private function cleanRsID($content)
    {
        /* From TBS script
         * Delete XML attributes relative to log of user modifications. Returns the number of deleted attributes.
        In order to insert such information, MsWord do split TBS tags with XML elements.
        After such attributes are deleted, we can concatenate duplicated XML elements. */

        $rs_lst = array('w:rsidR', 'w:rsidRPr');

        $nbr_del = 0;
        foreach ($rs_lst as $rs) {

            $rs_att = ' ' . $rs . '="';
            $rs_len = strlen($rs_att);

            $p = 0;
            while ($p !== false) {
                // search the attribute
                $ok = false;
                $p = strpos($content, $rs_att, $p);
                if ($p !== false) {
                    // attribute found, now seach tag bounds
                    $po = strpos($content, '<', $p);
                    $pc = strpos($content, '>', $p);
                    if (($pc !== false) && ($po !== false) && ($pc < $po)) { // means that the attribute is actually inside a tag
                        $p2 = strpos($content, '"', $p + $rs_len); // position of the delimiter that closes the attribute's value
                        if (($p2 !== false) && ($p2 < $pc)) {
                            // delete the attribute
                            $content = substr_replace($content, '', $p, $p2 - $p + 1);
                            $ok = true;
                            $nbr_del++;
                        }
                    }
                    if (!$ok) $p = $p + $rs_len;
                }
            }

        }

        // delete empty tags
        $content = str_replace('<w:rPr></w:rPr>', '', $content);
        $content = str_replace('<w:pPr></w:pPr>', '', $content);

        return $content;

    }

    private function cleanDuplicatedLayout($content)
    {
        // Return the number of deleted dublicates

        $wro = '<w:r';
        $wro_len = strlen($wro);

        $wrc = '</w:r';
        $wrc_len = strlen($wrc);

        $wto = '<w:t';
        $wto_len = strlen($wto);

        $wtc = '</w:t';
        $wtc_len = strlen($wtc);

        // number of replacements
        $nbr = 0;
        $wro_p = 0;
        while (($wro_p = $this->foundTag($content, $wro, $wro_p)) !== false) {
            $wto_p = $this->foundTag($content, $wto, $wro_p);
            if ($wto_p === false) {
                // throw new Exception('Error in the structure of the <w:r> element');
                continue;
            }
            $first = true;
            do {
                $ok = false;
                $wtc_p = $this->foundTag($content, $wtc, $wto_p);
                if ($wtc_p === false) {
                    // throw new Exception('Error in the structure of the <w:r> element');
                    continue;
                }
                $wrc_p = $this->foundTag($content, $wrc, $wro_p);
                if ($wrc_p === false) {
                    // throw new Exception('Error in the structure of the <w:r> element');
                    continue;
                }
                if (($wto_p < $wrc_p) && ($wtc_p < $wrc_p)) { // if the found <w:t> is actually included in the <w:r> element
                    $superflous = '';
                    $superflous_len = 0;
                    if ($first) {
                        $superflous = '</w:t></w:r>' . substr($content, $wro_p, ($wto_p + $wto_len) - $wro_p); // should be like: '</w:t></w:r><w:r>....<w:t'
                        $superflous_len = strlen($superflous);
                        $first = false;
                    }
                    $x = substr($content, $wtc_p + $superflous_len, 1);
                    if ((substr($content, $wtc_p, $superflous_len) === $superflous) && (($x === ' ') || ($x === '>'))) {
                        // if the <w:r> layout is the same same the next <w:r>, then we join it
                        $p_end = strpos($content, '>', $wtc_p + $superflous_len); //
                        if ($p_end === false) {
                            // throw new Exception("Error in the structure of the <w:t> tag");
                            continue;
                        }
                        $content = substr_replace($content, '', $wtc_p, $p_end - $wtc_p + 1);
                        $nbr++;
                        $ok = true;
                    }
                }
            } while ($ok);

            $wro_p = $wro_p + $wro_len;

        }

        return $content;

    }

    private function foundTag($Txt, $Tag, $PosBeg)
    {
        // Found the next tag of the asked type. (Not specific to MsWord, works for any XML)
        $len = strlen($Tag);
        $p = $PosBeg;
        while ($p !== false) {
            $p = strpos($Txt, $Tag, $p);
            if ($p === false) return false;
            $x = substr($Txt, $p + $len, 1);
            if (($x === ' ') || ($x === '/') || ($x === '>')) {
                return $p;
            } else {
                $p = $p + $len;
            }
        }
        return false;
    }

    private function cleanTag($content, array $tagList = [])
    {
        // Delete all tags of the types listed in the list. (Not specific to MsWord, works for any XML)
        foreach ($tagList as $tag) {
            $p = 0;
            while (($p = $this->foundTag($content, $tag, $p)) !== false) {
                // get the end of the tag
                $pe = strpos($content, '>', $p);
                if ($pe === false) {
                    // error in the XML formating
                    return false;
                }
                // delete the tag
                $content = substr_replace($content, '', $p, $pe - $p + 1);
            }
        }
        return $content;
    }

    /**
     * Scan directory for files starting with $file_name
     */
    private function getFilesStartingWith($file_name)
    {
        $dir = $this->tmpDir . "/word/";
        static $files;
        if (!isset($files)) {
            $files = scandir($dir);
        }

        $found_files = [];
        foreach ($files as $file) {
            if (is_file($dir . $file)) {
                if (strpos($file, $file_name) !== FALSE) {
                    $found_files[] = $file;
                }
            }
        }
        return $found_files;
    }
}
