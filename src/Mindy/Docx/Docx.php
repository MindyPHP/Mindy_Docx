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
}
