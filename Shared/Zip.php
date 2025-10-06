<?php
// namespace MyPhpWord\Shared;
class Zip
{
    /**
     * Archive filename (emulate ZipArchive::$filename).
     *
     * @var string
     */
    private $filename;
    /**
     * Temporary storage directory.
     *
     * @var string
     */
    private $tempDir;
    /**
     * Internal zip archive object.
     */
    private $zip;
    /**
     * Create new instance.
     */

    public function open($filename)
    {
        $result = true;
        $this->filename = $filename;
        $zip = new ZipArchive();

        if ($zip->open($this->filename) !== true) {
            $result = false;
        } else {
            $this->zip = $zip;
        }
        return $result;
    }

    /**
     * Close the active archive.
     *
     * @return bool
     */
    public function close()
    {
        try {
            $result = @$this->zip->close();
        } catch (Throwable $e) {
            $result = false;
        }
        if ($result === false) {
            throw new Exception("Could not close zip file {$this->filename}: ");
        }
        return true;
    }

    /**
     * Extract the archive contents (emulate \ZipArchive).
     *
     * @param string $destination
     *
     * @return bool
     *
     */
    public function extractTo($destination)
    {
        $this->tempDir = $destination;
        return $this->zip->extractTo($this->tempDir);
    }

    private static function delTree($dir)
    {
        $files = array_diff(scandir($dir), array('.', '..'));
        foreach ($files as $file) {
            (is_dir("$dir/$file")) ? self::delTree("$dir/$file") : unlink("$dir/$file");
        }
        return rmdir($dir);
    }

    public function create($destination)
    {
        if (file_exists($destination)) {
            unlink($destination);
        }

        if ($this->zip->open($destination, ZIPARCHIVE::CREATE)) {
            $source = $this->tempDir; 
            $source = str_replace('\\', '/', realpath($source)); 
            if (is_dir($source) === true) {
                $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($source), RecursiveIteratorIterator::SELF_FIRST);

                foreach ($files as $file) {
                    $file = str_replace('\\', '/', $file);

                    // Ignore "." and ".." folders
                    if (in_array(substr($file, strrpos($file, '/') + 1), array('.', '..'))) {
                        continue;
                    }

                    $file = realpath($file);
                    $file = str_replace('\\', '/', $file);

                    if (is_dir($file) === true) {
                        $this->zip->addEmptyDir(str_replace($source . '/', '', $file . '/'));
                    } else if (is_file($file) === true) {
                        $this->zip->addFromString(str_replace($source . '/', '', $file), file_get_contents($file));
                    }
                }
            } else if (is_file($source) === true) {
                $this->zip->addFromString(basename($source), file_get_contents($source));
            }
            $this->zip->close();

            //удалить распакованный word файл
            self::delTree(realpath($source));

        } else {
            return 'error';
        }
    }
}
