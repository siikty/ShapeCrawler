﻿using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class ShapeFillImage : IImage
{
    private readonly OpenXmlPart sdkOpenXmlPart;
    private readonly A.Blip aBlip;
    private ImagePart sdkImagePart;
    
    internal ShapeFillImage(OpenXmlPart sdkOpenXmlPart, A.BlipFill aBlipFill, ImagePart sdkImagePart)
    {
        this.sdkOpenXmlPart = sdkOpenXmlPart;
        this.aBlip = aBlipFill.Blip!;
        this.sdkImagePart = sdkImagePart;
    }

    public string MIME => this.sdkImagePart.ContentType;

    public string Name => Path.GetFileName(this.sdkImagePart.Uri.ToString());

    public void Update(Stream stream)
    {
        var isSharedImagePart = this.sdkOpenXmlPart.GetPartsOfType<ImagePart>().Count(x => x == this.sdkImagePart) > 1;
        if (isSharedImagePart)
        {
            var rId = $"rId-{Guid.NewGuid().ToString("N").Substring(0, 5)}";
            this.sdkImagePart = this.sdkOpenXmlPart.AddNewPart<ImagePart>("image/png", rId);
            this.aBlip.Embed!.Value = rId;
        }

        stream.Position = 0;
        this.sdkImagePart.FeedData(stream);
    }

    public void Update(byte[] bytes)
    {
        var stream = new MemoryStream(bytes);

        this.Update(stream);
    }

    public void Update(string file)
    {
        byte[] sourceBytes = File.ReadAllBytes(file);
        this.Update(sourceBytes);
    }

    public byte[] AsByteArray()
    {
        var stream = this.sdkImagePart.GetStream();
        var bytes = new byte[stream.Length];
        stream.Read(bytes, 0, (int)stream.Length);
        stream.Close();

        return bytes;
    }
}