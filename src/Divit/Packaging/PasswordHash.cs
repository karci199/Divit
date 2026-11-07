#nullable enable
namespace Divit.Packaging;

using System;
using System.Security.Cryptography;
using System.Text;

/// <summary>
/// The salted, iterated SHA-512 password hash Office uses for sheet, workbook and document protection
/// (ECMA-376 part 4, agile protection): H0 = SHA-512(salt + password as UTF-16LE), then H(n) = SHA-512(H(n-1) +
/// n as 32-bit little-endian), 100 000 times.
/// </summary>
/// <remarks>
/// Protection is not encryption: the file stays readable. The hash only decides whether Excel or Word lifts the
/// lock; with 100 000 rounds, guessing a password takes a long time per guess.
/// </remarks>
internal static class PasswordHash
{
    /// <summary>
    /// Office's own round count.
    /// </summary>
    public const int SpinCount = 100000;

    /// <summary>
    /// Office's algorithm name for the attribute.
    /// </summary>
    public const string Algorithm = "SHA-512";

    /// <summary>
    /// A 16-byte salt: random, or derived from <paramref name="seed"/> so the same document gives the same bytes. The
    /// seed must not hold the password: the salt is written into the file.
    /// </summary>
    /// <param name="seed">Text that makes the salt unique to this lock (part name, creation time); random when empty.</param>
    public static byte[] Salt(string? seed)
    {
        var salt = new byte[16];
        if (seed == null)
        {
            using (var random = RandomNumberGenerator.Create())
            {
                random.GetBytes(salt);
            }

            return salt;
        }

        using (var sha = SHA256.Create())
        {
            Array.Copy(sha.ComputeHash(Encoding.UTF8.GetBytes(seed)), salt, salt.Length);
        }

        return salt;
    }

    /// <summary>
    /// The hash of a password, Base64 as Office writes it.
    /// </summary>
    /// <param name="password">The password.</param>
    /// <param name="salt">The salt.</param>
    /// <param name="spinCount">Number of rounds after the first hash.</param>
    public static string Sha512(string password, byte[] salt, int spinCount = SpinCount)
    {
        var text = Encoding.Unicode.GetBytes(password);
        var first = new byte[salt.Length + text.Length];
        Array.Copy(salt, first, salt.Length);
        Array.Copy(text, 0, first, salt.Length, text.Length);

        // Every round hashes the previous hash and the round number: 64 + 4 bytes in one reused buffer.
        var buffer = new byte[68];
#if NET
        SHA512.HashData(first, buffer);
        for (var i = 0; i < spinCount; i++)
        {
            Counter(buffer, i);
            SHA512.HashData(buffer, buffer.AsSpan(0, 64));
        }
#else
        using (var sha = SHA512.Create())
        {
            Array.Copy(sha.ComputeHash(first), buffer, 64);
            for (var i = 0; i < spinCount; i++)
            {
                Counter(buffer, i);
                Array.Copy(sha.ComputeHash(buffer), buffer, 64);
            }
        }
#endif
        return Convert.ToBase64String(buffer, 0, 64);
    }

    private static void Counter(byte[] buffer, int round)
    {
        buffer[64] = (byte)round;
        buffer[65] = (byte)(round >> 8);
        buffer[66] = (byte)(round >> 16);
        buffer[67] = (byte)(round >> 24);
    }
}
