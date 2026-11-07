#nullable enable
namespace Divit.Pdf;

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

/// <summary>
/// A growing buffer of the ASCII bytes of PDF syntax: operators, names, numbers, hex strings. Numbers are written
/// digit by digit, without the runtime's formatting, so both targets give the same bytes.
/// </summary>
internal sealed class PdfBuffer
{
    private byte[] _data;

    public PdfBuffer(int capacity = 4096)
    {
        _data = new byte[capacity];
    }

    public int Length { get; private set; }

    public byte[] Data => _data;

    public PdfBuffer Append(string ascii)
    {
        Ensure(ascii.Length);
        for (var i = 0; i < ascii.Length; i++)
        {
            _data[Length++] = (byte)ascii[i];
        }

        return this;
    }

    public PdfBuffer Append(char c)
    {
        Ensure(1);
        _data[Length++] = (byte)c;
        return this;
    }

    public PdfBuffer Append(byte[] bytes, int count)
    {
        Ensure(count);
        Buffer.BlockCopy(bytes, 0, _data, Length, count);
        Length += count;
        return this;
    }

    /// <summary>
    /// A number rounded to three decimals, without exponent or trailing zeros, then a space.
    /// </summary>
#if NET
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
    public PdfBuffer Number(double value)
    {
        Ensure(24);
        var scaled = (long)Math.Round(value * 1000, MidpointRounding.AwayFromZero);
        if (scaled < 0)
        {
            _data[Length++] = (byte)'-';
            scaled = -scaled;
        }

        var whole = scaled / 1000;
        var fraction = (int)(scaled % 1000);
        Integer(whole);
        if (fraction != 0)
        {
            _data[Length++] = (byte)'.';
            var digits = fraction % 10 != 0 ? 3 : fraction % 100 != 0 ? 2 : 1;
            var divisor = 100;
            for (var i = 0; i < digits; i++)
            {
                _data[Length++] = (byte)('0' + (fraction / divisor % 10));
                divisor /= 10;
            }
        }

        _data[Length++] = (byte)' ';
        return this;
    }

    /// <summary>
    /// A whole number in decimal.
    /// </summary>
#if NET
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
    public PdfBuffer Integer(long value)
    {
        Ensure(21);
        if (value < 0)
        {
            _data[Length++] = (byte)'-';
            value = -value;
        }

        var start = Length;
        do
        {
            _data[Length++] = (byte)('0' + (value % 10));
            value /= 10;
        }
        while (value > 0);

        Array.Reverse(_data, start, Length - start);
        return this;
    }

    /// <summary>
    /// Four upper-case hex digits: one glyph id of a text string.
    /// </summary>
#if NET
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.AggressiveOptimization)]
#endif
    public void Hex4(int value)
    {
        Ensure(4);
        for (var shift = 12; shift >= 0; shift -= 4)
        {
            var digit = (value >> shift) & 0xF;
            _data[Length++] = (byte)(digit < 10 ? '0' + digit : 'A' + digit - 10);
        }
    }

    public void Clear() => Length = 0;

    private void Ensure(int more)
    {
        if (Length + more > _data.Length)
        {
            Array.Resize(ref _data, Math.Max(_data.Length * 2, Length + more));
        }
    }
}

/// <summary>
/// Writes the objects of a PDF file one after another, keeping the offset of each for the cross-reference table.
/// Object ids are reserved first, so objects can point at ones written later.
/// </summary>
internal sealed class PdfFile
{
    private readonly Stream _stream;
    private readonly PdfBuffer _buffer = new PdfBuffer(1 << 16);
    private readonly List<long> _offsets = new List<long> { 0 };
    private long _position;

    public PdfFile(Stream stream)
    {
        _stream = stream;

        // The comment after the version holds bytes above 127, so tools treat the file as binary.
        _buffer.Append("%PDF-1.7\n%").Append(new byte[] { 0xE2, 0xE3, 0xCF, 0xD3, 0x0A }, 5);
        Flush();
    }

    /// <summary>
    /// Takes the next object id; the object is written later with <see cref="Begin"/>.
    /// </summary>
    public int Reserve()
    {
        _offsets.Add(-1);
        return _offsets.Count - 1;
    }

    /// <summary>
    /// Starts object <paramref name="id"/>; its dictionary or value goes into the returned buffer.
    /// </summary>
    public PdfBuffer Begin(int id)
    {
        _offsets[id] = _position + _buffer.Length;
        return _buffer.Integer(id).Append(" 0 obj\n");
    }

    public void End()
    {
        _buffer.Append("\nendobj\n");
        if (_buffer.Length > 1 << 15)
        {
            Flush();
        }
    }

    /// <summary>
    /// A stream object: its dictionary entries (without the length and filter), then the data, compressed unless
    /// it already is.
    /// </summary>
    public void Stream(int id, string entries, byte[] data, int count, bool compress = true)
    {
        byte[] packed = data;
        var length = count;
        if (compress)
        {
            packed = Deflate(data, count);
            length = packed.Length;
        }

        var b = Begin(id).Append("<< ").Append(entries).Append(compress ? "/Filter /FlateDecode " : string.Empty).Append("/Length ").Integer(length).Append(" >>\nstream\n");
        Flush();
        _stream.Write(packed, 0, length);
        _position += length;
        b.Append("\nendstream");
        End();
    }

    /// <summary>
    /// The cross-reference table and the trailer; the file is complete after it.
    /// </summary>
    public void Close(int catalog, int info)
    {
        var xref = _position + _buffer.Length;
        _buffer.Append("xref\n0 ").Integer(_offsets.Count).Append("\n0000000000 65535 f \n");
        for (var id = 1; id < _offsets.Count; id++)
        {
            var offset = _offsets[id].ToString("D10", System.Globalization.CultureInfo.InvariantCulture);
            _buffer.Append(offset).Append(" 00000 n \n");
        }

        _buffer.Append("trailer\n<< /Size ").Integer(_offsets.Count).Append(" /Root ").Integer(catalog).Append(" 0 R /Info ").Integer(info)
            .Append(" 0 R >>\nstartxref\n").Integer(xref).Append("\n%%EOF\n");
        Flush();
    }

    /// <summary>
    /// Data in the zlib format that FlateDecode reads: a two-byte header, the deflated bytes, the Adler-32 sum.
    /// </summary>
    public static byte[] Deflate(byte[] data, int count)
    {
        using (var packed = new MemoryStream(count / 2 + 64))
        {
            packed.WriteByte(0x78);
            packed.WriteByte(0x9C);
            using (var deflate = new DeflateStream(packed, CompressionLevel.Optimal, true))
            {
                deflate.Write(data, 0, count);
            }

            uint a = 1, b = 0;
            for (var i = 0; i < count;)
            {
                // Sums stay below 2^32 for 5552 bytes between reductions.
                var end = Math.Min(count, i + 5552);
                for (; i < end; i++)
                {
                    a += data[i];
                    b += a;
                }

                a %= 65521;
                b %= 65521;
            }

            packed.WriteByte((byte)(b >> 8));
            packed.WriteByte((byte)b);
            packed.WriteByte((byte)(a >> 8));
            packed.WriteByte((byte)a);
            return packed.ToArray();
        }
    }

    private void Flush()
    {
        _stream.Write(_buffer.Data, 0, _buffer.Length);
        _position += _buffer.Length;
        _buffer.Clear();
    }
}
