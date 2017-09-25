# CaptureCodePlexDiscussions
Capture all the discussions for a CodePlex project

CodePlex is shutting down on December 15th 2017. As of the current plan of record (see Brian Harry's blog post at https://blogs.msdn.microsoft.com/bharry/2017/03/31/shutting-down-codeplex/), all discussion threads will be discarded at this point, although sources and downloads will be preserved.

There have been some complaints in the comments on the blog page referenced above about the loss of discussions. However, the last post on the subject was the following entry from Alex Mullans of Microsoft four months ago:

@dahmaninator We’ll be starting that work in the next sprint (which starts on 6/19), so I’d expect it to be available ~3 weeks after that.

However, absolutely nothing has been heard of this since. Being unprepared to lose the discussions from my own projects, I've written my own discussion capture program, provided here.

This project provides a program that will download all the discussion content for a CodePlex project and save it as a Word document (for easy reading), or an XML file (for programmatic access), or both.

The program screen scrapes the CodePlex site, using WatiN to drive Internet Explorer. It is not particularly quick, but it appears to be fairly thorough: I've included some sample results of captures from different projects.
