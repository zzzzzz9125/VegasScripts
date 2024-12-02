# VegasScripts
Other Vegas Pro scripts written by me.

### PsdLayerSplitter
A simple script to split layers of a .psd file in Vegas with one-click.

Note that If you hold down Ctrl and click the script icon on the toolbar, the selected event will be converted to Stream 1 (in a programming sense, it's a video stream with Index 0).

### Add Note Name to File
A simple script to add note names to audio files and display them on events.

![20241201223326](https://github.com/user-attachments/assets/7b160772-e00a-41e7-b77b-b6e6604d45c7)
![20241201221820](https://github.com/user-attachments/assets/61f76345-ec7a-4387-8346-f38830d88937)

This script supports almost all types of media files (WAV, FLAC, MP3, OGG, MP4, etc.). For WAV, FLAC and MP3, the script generates new media files to replace them. For other formats, it can add `.sfl` files to save note name information.

Run the script after selecting the events on the timeline. Once you've clicked OK, you'll need to manually update the note name info, which you can do in one of three ways: 
1. Perform any event paste operation in the project;
2. Drag the files into the timeline again;
3. Reopen the project.
