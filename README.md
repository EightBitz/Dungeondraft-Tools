<!DOCTYPE html>
<html>
<head>
<title>EIGHTBITZ'S DUNGEONDRAFT TOOLS README</title>
</head>
<body>
<h1>EightBitz's Dungeondraft Tools</h1>
<p>What follows is a brief description of what this is and what it does, but please read the full documentation as it contains important details.</p>

<h2>Getting Started</h2>

<p>To use this program:</p>
<ol>
<li><a href="https://github.com/EightBitz/Dungeondraft-Tools/archive/Version-3.0.zip">Download the Zip archive</a>.</li>
<li>Extract the folder.</li>
<li>Run Setup.exe</li>
<li>Double click the "DDTools" desktop icon.</li>
</ol>

<p>
If you want to convert files to .webp format, you will also need to download and install <a href="https://imagemagick.org/script/download.php#windows">ImageMagick for Windows</a>.

<h2>What Does it Do?</h2>

This program offers six tools to facilitate creating and managing custom asset packs for Dungeondraft.
<ul>
<li>Tag assets</li>
<li>Convert Assets</li>
<li>Convert Packs</li>
<li>Copy Assets</li>
<li>Data Files</li>
<li>Pack Assets</li>
<li>Unpack Assets</li>
</ul>

<p>
If you're looking for my Custom Tags tool, <a href="https://github.com/EightBitz/Dungeondraft-Custom-Tags">that can be found here</a>.

<h3>Tag Assets</h3>

<p>For people who want to create their own custom asset packs, the Tag Assets tool will create a properly formatted default.dungeondraft_tags file. This is the file that Dungeondraft uses to determine which assets in the object library are associated with which tags.</p>

<h3>Convert Assets</h3>

<p>The Convert Assets tool will convert supported image formats in selected folders to .webp format. .webp offers lossless compression, so the files on disk end up being smaller. That being said, this will not save resources in Dungeondraft, as when .webp files are loaded into memory, they are uncompressed. I added this function at the suggestion of someone on the Dungeondraft Discord server. Whether or not you find it ultimately beneficial, I will leave to you.</p>

<h3>Convert Packs</h3>

<p>This works similar to the Convert Assets tool, except it works on files that are already packed. This tool calls dungeondraft-unpack.exe to unpack selected files, then it calls ImageMagick to convert assets to webp, then it repacks the converted asset folders.</p>

<h3>Copy Assets</h3>

<p>This tool is primarily designed to copy Campaign Cartographer 3 Plus (CC3+) assets into a textures\objects and/or textures\portals folder structure for Dungeondraft. This tool will work for other assets as well, but Campaign Cartographer has two important naming conventions that this tool uses to manage how it copies assets.</p>

<h3>Copy Tiles</h3>

<p>This tool is similar to Copy Assets, except that it allows you to copy assets as patterns, terrain or tilesets.</p>

<h3>Data Files</h3>

<p>This tool allows you to easily create and manage data files for tilesets and walls.</p>

<h3>Map Details</h3>

<p>This tool provides details for selected maps, such as when they were created, which version of Dungeondraft they were created with, which cusom asset packs they use, and which assets they use from those asset packs.</p>

<h3>Pack Assets</h3>

<p>This tool will create .dungeondraft_pack files from selected asset folders.</p>

<h3>Unpack Assets</h3>

This tool will unpack .dungeondraft_pack files.
</body>
</html> 
