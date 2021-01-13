# FastDirectoryEnumerator-VB.NET
 VB.NET version of FastDirectoryEnumerator by wilsone8

### Original Description

A Faster Directory Enumerator
Author:  wilsone8
Date:    2009 Aug 27
License: CPOL
Describes how to create a significantly faster enumerator for the attributes of all the files in a directory.

## Introduction

The .NET Framework's Directory class includes methods for querying the list of files in a directory. For each file, you can also then query the attributes of the file, such as the size, creation date, etc. However, when querying files on a remote PC, this can be very inefficient; a potentially expensive network round-trip is needed to retrieve each file's attributes. This article describes a much more efficient implementation that is approximately 3x faster.

## Background

Let's assume you are writing an application that needs to find the most recently modified file in a directory. To implement this, you might have a function similar to the following:

    DateTime GetLastFileModifiedSlow(string dir)
    {
        DateTime retval = DateTime.MinValue;
        
        string [] files = Directory.GetFiles(dir);
        for (int i=0; i<files.Length; i++)
        {
            DateTime lastWriteTime = File.GetLastWriteTime(files[i]);
            if (lastWriteTime > retval)
            {
                retval = lastWriteTime;
            }
        }
        
        return retval;
    }

That function certainly works, but it suffers from some very poor performance characteristics:

  -  GetFiles must allocate a potentially very large array.
  -  GetFiles must wait for the entire directory's entries to be returned before returning.
  -  For each file, a potentially expensive query is sent to the file system. No attempt is made to perform any sort of batch query.

You might think that converting to DirectoryInfo.GetFileSystemInfos would improve item #3:

    DateTime GetLastFileModifiedSlow2(string dir)
    {
        DateTime retval = DateTime.MinValue;
        
        DirectoryInfo dirInfo = new DirectoryInfo(dir);

        FileInfo[] files = dirInfo.GetFiles();
        for (int i=0; i<files.Length; i++)
        {
            if (files[i].LastWriteTime > retval)
            {
                retval = lastWriteTime;
            }
        }
        
        return retval;
    }

This doesn't change anything however: the objects returned by GetFiles() are not initialized with any data, and will all query the file system the first time any property is accessed.

## Making it Faster

The attached test application includes the FastDirectoryEnumerator class in FastDirectoryEnumerator.cs. Using the GetFiles method, we can write the equivalent of our first slow method.

    DateTime GetLastFileModifiedFast(string dir)
    {
        DateTime retval = DateTime.MinValue;
        
        FileData [] files = FastDirectoryEnumerator.GetFiles(dir);
        for (int i=0; i<files.Length; i++)
        {
            if (files[i].LastWriteTime > retval)
            {
                retval = lastWriteTime;
            }
        }
        
        return retval;
    }

The FileData object provides all the standard attributes for a file that the FileInfo class provides.

## Making it Even Faster

Use one of the overloads of the EnumerateFiles method to enumerate over all the files in a directory. The enumeration returns a FileData object.

Below is an example of the same method using FastDirectoryEnumerator:

    DateTime GetLastFileModifiedFast(string dir)
    {
        DateTime retval = DateTime.MinValue;

        foreach (FileData f in FastDirectoryEnumerator.EnumerateFiles(dir))
        {
            if (f.LastWriteTime > retval)
            {
                retval = f.LastWriteTime;
            }
        }

        return retval;
    }

## Performance

The test application allows you to create a large number of files in a directory, then test the time it takes to enumerate using all three methods. I used a directory with 3000 files and ran each test three times to give the best answer possible for each test.

Using a path on my local hard drive resulted in the following times:

  -  Directory.GetFiles method: ~225ms
  -  DirectoryInfo.GetFiles method: ~230ms
  -  FastDirectoryEnumerator.GetFiles method: ~33ms
  -  FastDirectoryEnumerator.EnumerateFiles method: ~27ms

That is roughly a 8.5x increase in performance between the fastest and the slowest methods. The performance is even more pronounced when the files are on a UNC path. For this test, I used the same directory as the previous test. The only difference is that I referenced the directory by a UNC share name instead of the local path. At the time of the test, I was connected to my home wireless network.

  -  Directory.GetFiles method: ~43,860ms
  -  DirectoryInfo.GetFiles method: ~44,000ms
  -  FastDirectoryEnumerator.GetFiles method: ~55ms
  -  FastDirectoryEnumerator.EnumerateFiles method: ~53ms

That is roughly a 830x increase in performance, and more than 2 orders of magnitude! And, the gap only increases as the latency to the PC containing the files increases.

## Why is it Faster?

As mentioned above, Directory.GetFiles and DirectoryInfo.GetFiles have a number of disadvantages. The most significant is that they throw away information and do not efficiently allow you to retrieve information about multiple files at the same time.

Internally, Directory.GetFiles is implemented as a wrapper over the Win32 FindFirstFile/FindNextFile functions. These functions all return information about each file that is enumerated that the GetFiles() method throws away when it returns the file names. They also retrieve information about multiple files with a single network message.

The FastDirectoryEnumerator keeps this information and returns it in the FileData class. This substantially reduces the number of network round-trips needed to accomplish the same task.

## History

  -  8-13-2009: Initial version.
  -  8-14-2009: Added security checks, parameter checking, and the GetFiles method.
  -  8-24-2009: Fixed the AllDirectories search using GetFiles. Removed note about .NET 4.0 including something similar.
  -  9-08-2009: Fixed the AllDirectories search when filter is not * or *.*.

## License

This article, along with any associated source code and files, is licensed under [The Code Project Open License (CPOL)](http://www.codeproject.com/info/cpol10.aspx)
