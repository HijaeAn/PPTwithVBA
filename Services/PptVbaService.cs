using PptWithVba.Models;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace PptWithVba.Services;

/// <summary>
/// C# → VBA 방식 PPT 자동화 서비스.
/// 흐름: XML 데이터 파일 생성 → PowerPoint COM 실행 → VBA 모듈 주입 → 매크로 실행 → .pptx 저장
/// </summary>
public sealed class PptVbaService
{
    private const string ModuleName = "SlideCreator";
    private const string MacroName  = "SlideCreator.CreateSlides";

    // ─── VBA 코드 (C# 런타임에 PowerPoint VBProject에 주입됨) ──────────────
    private const string VbaCode = """
        Option Explicit

        Sub CreateSlides(dataFilePath As String, outputPath As String)
            Dim xmlDoc As Object
            Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
            xmlDoc.async = False

            If Not xmlDoc.Load(dataFilePath) Then
                Err.Raise 1001, "CreateSlides", "XML 로드 실패: " & xmlDoc.parseError.reason
            End If

            Dim prs As Presentation
            Set prs = ActivePresentation

            Dim i As Integer
            For i = prs.Slides.Count To 1 Step -1
                prs.Slides(i).Delete
            Next i

            Dim slideNodes As Object
            Set slideNodes = xmlDoc.SelectNodes("//Slide")

            Dim idx As Integer: idx = 1
            Dim node As Object
            For Each node In slideNodes
                BuildSlide prs, node, idx
                idx = idx + 1
            Next node

            prs.SaveAs outputPath, ppSaveAsOpenXMLPresentation
        End Sub

        Private Sub BuildSlide(prs As Presentation, slideNode As Object, idx As Integer)
            Const MARGIN   As Single = 30
            Const TITLE_H  As Single = 55
            Const LINE_GAP As Single = 10

            Dim sld As Slide
            Set sld = prs.Slides.Add(idx, ppLayoutBlank)

            Dim slideW As Single: slideW = prs.PageSetup.SlideWidth
            Dim slideH As Single: slideH = prs.PageSetup.SlideHeight

            ' 흰색 배경
            sld.Background.Fill.Solid
            sld.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)

            ' ── 제목 ──────────────────────────────────────────────────────
            Dim nd As Object
            Dim titleText As String
            Set nd = slideNode.SelectSingleNode("Title")
            If Not nd Is Nothing Then titleText = nd.Text

            Dim titleBox As Shape
            Set titleBox = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                MARGIN, MARGIN, slideW - MARGIN * 2, TITLE_H)
            With titleBox.TextFrame
                .TextRange.Text = titleText
                .TextRange.Font.Name = "맑은 고딕"
                .TextRange.Font.Size = 28
                .TextRange.Font.Bold = True
                .TextRange.Font.Color.RGB = RGB(31, 73, 125)
                .VerticalAnchor = msoAnchorMiddle
                .WordWrap = msoTrue
            End With

            ' ── 제목 하단 구분선 ──────────────────────────────────────────
            Dim lineY As Single: lineY = MARGIN + TITLE_H + 4
            With sld.Shapes.AddLine(MARGIN, lineY, slideW - MARGIN, lineY).Line
                .ForeColor.RGB = RGB(31, 73, 125)
                .Weight = 1.5
            End With

            ' ── 콘텐츠 영역 기준값 ────────────────────────────────────────
            Dim contentTop As Single: contentTop = lineY + LINE_GAP
            Dim contentH   As Single: contentH   = slideH - contentTop - MARGIN

            ' ── 이미지 (왼쪽 45%, 원본 비율 유지) ────────────────────────
            Dim imgPath As String
            Set nd = slideNode.SelectSingleNode("ImagePath")
            If Not nd Is Nothing Then imgPath = Trim(nd.Text)

            Dim imgAreaW As Single: imgAreaW = 0
            If imgPath <> "" And Dir(imgPath) <> "" Then
                Dim maxW As Single: maxW = (slideW - MARGIN * 2) * 0.45
                Dim pic As Shape
                Set pic = sld.Shapes.AddPicture(imgPath, msoFalse, msoCTrue, _
                    MARGIN, contentTop, maxW, contentH)

                Dim r As Single: r = pic.Width / pic.Height
                If pic.Height > contentH Then
                    pic.Height = contentH: pic.Width = contentH * r
                End If
                If pic.Width > maxW Then
                    pic.Width = maxW: pic.Height = maxW / r
                End If
                imgAreaW = pic.Width
            End If

            ' ── 내용 텍스트 박스 (이미지 있으면 오른쪽, 없으면 전체 폭) ──
            Dim items As Object
            Set items = slideNode.SelectNodes("Content/Item")
            If items.Length = 0 Then Exit Sub

            Dim textLeft As Single, textW As Single
            If imgAreaW > 0 Then
                textLeft = MARGIN + imgAreaW + 15
                textW    = slideW - textLeft - MARGIN
            Else
                textLeft = MARGIN
                textW    = slideW - MARGIN * 2
            End If

            Dim buf As String: buf = ""
            Dim item As Object
            For Each item In items
                If buf <> "" Then buf = buf & vbCr
                buf = buf & Chr(8226) & "  " & item.Text
            Next item

            Dim cb As Shape
            Set cb = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                textLeft, contentTop, textW, contentH)
            With cb.TextFrame
                .TextRange.Text = buf
                .TextRange.Font.Name = "맑은 고딕"
                .TextRange.Font.Size = 16
                .TextRange.Font.Color.RGB = RGB(50, 50, 50)
                .WordWrap = msoTrue
                .AutoSize = ppAutoSizeNone
            End With
        End Sub
        """;

    // ─── 공개 API ───────────────────────────────────────────────────────────

    /// <summary>
    /// 슬라이드 목록을 받아 .pptx 파일을 생성한다.
    /// PowerPoint가 설치되어 있어야 하며, Trust Center의
    /// "VBA 프로젝트 개체 모델 액세스 신뢰" 설정이 활성화되어 있어야 한다.
    /// </summary>
    public void CreatePresentation(IList<SlideContent> slides, string outputPath)
    {
        ArgumentNullException.ThrowIfNull(slides);
        if (slides.Count == 0)
            throw new ArgumentException("슬라이드가 없습니다.", nameof(slides));

        string dataFile = WriteDataXml(slides);
        object? comApp = null;
        object? comPrs = null;

        try
        {
            var pptType = Type.GetTypeFromProgID("PowerPoint.Application")
                ?? throw new InvalidOperationException("PowerPoint가 설치되어 있지 않습니다.");

            comApp = Activator.CreateInstance(pptType)!;
            dynamic app = comApp;
            app.Visible = true;                    // -1 = msoTrue

            comPrs = app.Presentations.Add(false); // 0  = msoFalse (창 없이 추가)
            dynamic prs = comPrs;

            InjectVbaModule(prs);

            app.Run(MacroName, dataFile, outputPath);
        }
        catch (COMException ex) when (
            ex.Message.Contains("Programmatic", StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("trust",        StringComparison.OrdinalIgnoreCase) ||
            ex.Message.Contains("신뢰",          StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "VBA 프로젝트 개체 모델 접근이 차단되어 있습니다.\n" +
                "PowerPoint → 파일 → 옵션 → 보안 센터 → 보안 센터 설정\n" +
                "→ 매크로 설정 → \"VBA 프로젝트 개체 모델에 대한 액세스 신뢰\" 를 체크하세요.", ex);
        }
        finally
        {
            File.Delete(dataFile);
            try { ((dynamic?)comPrs)?.Close();   } catch { }
            try { ((dynamic?)comApp)?.Quit();    } catch { }
            if (comPrs is not null) Marshal.ReleaseComObject(comPrs);
            if (comApp is not null) Marshal.ReleaseComObject(comApp);
        }
    }

    // ─── 내부 헬퍼 ─────────────────────────────────────────────────────────

    private static void InjectVbaModule(dynamic prs)
    {
        dynamic vbProject = prs.VBProject;
        dynamic vbComp    = vbProject.VBComponents.Add(1); // 1 = vbext_ct_StdModule
        vbComp.Name = ModuleName;
        vbComp.CodeModule.AddFromString(VbaCode);
    }

    private static string WriteDataXml(IList<SlideContent> slides)
    {
        string path = Path.Combine(Path.GetTempPath(), $"ppt_{Guid.NewGuid():N}.xml");

        new XDocument(
            new XElement("Slides",
                slides.Select(s => new XElement("Slide",
                    new XElement("Title",     s.Title),
                    new XElement("ImagePath", s.ImagePath ?? string.Empty),
                    new XElement("Content",
                        s.Content.Select(c => new XElement("Item", c)))
                ))
            )
        ).Save(path);

        return path;
    }
}
