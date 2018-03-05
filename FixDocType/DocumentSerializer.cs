using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace FixDocType
{
    public class DocumentSerializer : IDisposable
    {
        private const string UniquePostfix = "STEWIESAYSWHATTHEDEUCE";
        private readonly MemoryStream _docStream = null;
        private readonly WordprocessingDocument _wordDocument = null;
        private string _documentText = null;

        public DocumentSerializer(byte[] document)
        {
            _docStream = new MemoryStream(document, true);
            _wordDocument = WordprocessingDocument.Open(_docStream, true);
            using (var reader = new StreamReader(_wordDocument.MainDocumentPart.GetStream()))
            {
                _documentText = reader.ReadToEnd();
            }
        }

        const string WrongPattern = @"DOCPROPERTY CrmProperty(?<junk1>.*?>)(?<index>[0-9]{1,3}\s)(?<junk2>.*?)(?<ending>\\\* CHARFORMAT\s+\\\*\s+MERGEFORMAT)";
        public void Fix()
        {
            var match = Regex.Match(_documentText, WrongPattern);
            while (match.Success)
            {
                var junk1 = match.Groups["junk1"];
                var index = match.Groups["index"];
                var junk2 = match.Groups["junk2"];
                var ending = match.Groups["ending"];

                var fixedValue = "DOCPROPERTY CRMPROPERTY" + index + ending;
                _documentText = _documentText.Replace(match.Value, fixedValue);

                match = Regex.Match(_documentText, WrongPattern);
            }
        }

        public void Dispose()
        {
            try
            {
                if (_wordDocument != null)
                {
                    _wordDocument.Dispose();
                    _documentText = null;
                }

                if (_docStream != null)
                {
                    _docStream.Dispose();
                }
            }
            catch (Exception ex)
            {
                Program.Error(ex.ToString());
            }
        }

        //public void Replace(string find, string replace)
        //{
        //    var regexPatternFormat = string.Format("{{0}}(?!{0}|\\d)", UniquePostfix);
        //    var pattern = string.Format(regexPatternFormat, find);
        //    var newTempKey = replace + UniquePostfix;
        //    _documentText = Regex.Replace(_documentText, pattern, newTempKey);
        //}

        public byte[] ToBytes()
        {
            _documentText = _documentText.Replace(UniquePostfix, string.Empty);
            using (var writer = new StreamWriter(_wordDocument.MainDocumentPart.GetStream(FileMode.Create)))
            {
                writer.Write(_documentText);
            }

            return _docStream.ToArray();
        }
    }
}
