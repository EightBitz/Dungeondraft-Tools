# EightBitz's Dungeondraft Tools

What follows is a brief description of what this is and what it does, but please read [the full documentation](https://github.com/EightBitz/Dungeondraft-Tools/blob/Version-2.5/EightBitz's%20Dungeondraft%20Tools%20Documentation.pdf) as it contains important details.

## Getting Started

To use this program, download the following files from this repository, and put them all in the same folder:

* [DDTools.exe](https://github.com/EightBitz/Dungeondraft-Tools/blob/Version-2.5/DDTools.exe)
* [Newtonsoft.Json.dll](https://github.com/EightBitz/Dungeondraft-Tools/blob/Version-2.5/Newtonsoft.Json.dll)
* [dungeondraft-pack.exe](https://github.com/EightBitz/Dungeondraft-Tools/blob/Version-2.5/dungeondraft-pack.exe)
* [dungeondraft-unpack.exe](https://github.com/EightBitz/Dungeondraft-Tools/blob/Version-2.5/dungeondraft-unpack.exe)

If you want to convert files to .webp format, you will also need to download and install [ImageMagick for Windows](https://imagemagick.org/script/download.php#windows).

You don't need anything from the Source folder unless you want to peek under the hood and look at the code. For those who do, this was written in Visual Basic .NET using Visual Studio 2019, Community Edition.

## What Does it Do?

This program offers six tools to facilitate creating and managing custom asset packs for Dungeondraft.

* Tag assets
* Convert Assets
* Convert Packs
* Copy Assets
* Data Files
* Pack Assets
* Unpack Assets

### Tag Assets

For people who want to create their own custom asset packs, the Tag Assets tool will create a properly
formatted default.dungeondraft_tags file. This is the file that Dungeondraft uses to determine which
assets in the object library are associated with which tags.

### Convert Assets

The Convert Assets tool will convert supported image formats in selected folders to .webp format.
.webp offers lossless compression, so the files on disk end up being smaller. That being said, this will not
save resources in Dungeondraft, as when .webp files are loaded into memory, they are uncompressed.
I added this function at the suggestion of someone on the Dungeondraft Discord server. Whether or not
you find it ultimately beneficial, I will leave to you.

### Convert Packs

This works similar to the Convert Assets tool, except it works on files that are already packed. This tool
calls dungeondraft-unpack.exe to unpack selected files, then it calls ImageMagick to convert assets to
webp, then it repacks the converted asset folders.

### Copy Assets

This tool is primarily designed to copy Campaign Cartographer 3 Plus (CC3+) assets into a
textures\objects and/or textures\portals folder structure for Dungeondraft. This tool will work for other
assets as well, but Campaign Cartographer has two important naming conventions that this tool uses to
manage how it copies assets.

### Copy Tiles
This tool is similar to Copy Assets, except that it allows you to copy assets as patterns, terrain or tilesets.

### Data Files
This tool allows you to easily create and manage data files for tilesets and walls.

### Map Details
This tool provides details for selected maps, such as when they were created, which version of Dungeondraft they were created with, which cusom asset packs they use, and which assets they use from those asset packs.

### Pack Assets

This tool will create .dungeondraft_pack files from selected asset folders.

### Unpack Assets

This tool will unpack .dungeondraft_pack files.
