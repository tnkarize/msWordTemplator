 <?php
class ManageDocument
{
    public $Document_Name;
    public $Document_Type;
    public $Document_Path;
    public $Document_Extension;
    public $tmppath = 'tmp/document';
    public function __construct($name, $type, $path, $extension)
    {
        $this->Document_Name = $name;
        $this->Document_Type = $type;
        $this->Document_Path = $path;
        $this->Document_Extension = $extension;
    }
    public function edit($c, $e)
    {
        if ($this->Document_Type == 'Docx' || $this->Document_Type == 'Microsoft Word' || $this->Document_Type == 'Word')
        {
        rename($this->Document_Path.'/'.$this->Document_Name.'.'.$this->Document_Extension, $this->Document_Path.'/'.$this->Document_Name.'.zip');
        $path = $this->Document_Path.'/'.$this->Document_Name.'.zip';
        $zip = new ZipArchive;
        $res = $zip->open($path);
        if($res === true)
        {
         $zip->extractTo($this->Document_Path.'/'.$this->Document_Name);
        $zip->close();
         unlink($path);
        }
        $fle = file_get_contents($this->Document_Path.'/'.$this->Document_Name.'/word/document.xml');
        for ($i=0; $i<count($c); $i++)
        {
            $fle = str_replace($c[$i], $e[$i], $fle);
        }
        file_put_contents($this->Document_Path.'/'.$this->Document_Name.'/word/document.xml', $fle);
        $rootPath = realpath($this->Document_Path.'/'.$this->Document_Name.'/');
         $zp = new ZipArchive;
$zp->open($this->Document_Name.'.zip', ZipArchive::CREATE | ZIPARCHIVE::OVERWRITE);
// Create recursive directory iterator
/** @var SplFileInfo[] $files */
$files = new RecursiveIteratorIterator(
    new RecursiveDirectoryIterator($rootPath),
    RecursiveIteratorIterator::LEAVES_ONLY
);
foreach ($files as $name => $file)
{
    // Skip directories (they would be added automatically)
    if (!$file->isDir())
    {
        // Get real and relative path for current file
        $filePath = $file->getRealPath();
        $relativePath = substr($filePath, strlen($rootPath) + 1);
        // Add current file to archive
        $zp->addFile($filePath, $relativePath);
    }
}
// Zip archive will be created only after closing object
$zp->close();
        }
    }
}
?>
