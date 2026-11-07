#nullable enable
namespace Divit.Packaging;

using System;
using System.IO;

/// <summary>
/// Saves a file so an existing one is replaced only once the new one is complete: written to a temporary file next
/// to it, then moved over it. Every format saves through it.
/// </summary>
internal static class AtomicFile
{
    /// <summary>
    /// Writes the file through <paramref name="write"/>; on failure the old file stays and the temporary one goes.
    /// </summary>
    /// <param name="path">Target file.</param>
    /// <param name="write">Writes the content into the stream.</param>
    public static void Save(string path, Action<Stream> write)
    {
        var temp = path + ".tmp";
        try
        {
            using (var file = new FileStream(temp, FileMode.Create, FileAccess.Write, FileShare.None, 65536))
            {
                write(file);
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            File.Move(temp, path);
        }
        finally
        {
            if (File.Exists(temp))
            {
                File.Delete(temp);
            }
        }
    }
}
