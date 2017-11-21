using System.IO;

namespace UnlockVBA
{
    internal class MicrosoftOffice
    {
        public MicrosoftOffice(string fileIn, string fileOut) : this()
        {
            FileIn = fileIn;
            FileOut = fileOut;
        }

        public MicrosoftOffice()
        {
            _sourceSeq = new byte[] { 0x44, 0x50, 0x42 };
            _targetSeq = new byte[] { 0x44, 0x50, 0x78 };
        }

        public string FileIn { get; }
        public string FileOut { get; }
        private readonly byte[] _sourceSeq;

        private readonly byte[] _targetSeq;

        public void Decode()
        {
            if (File.Exists(FileOut))
            {
                File.Delete(FileOut);
            }

            FileStream sourceStream = File.OpenRead(FileIn);
            FileStream targetStream = File.Create(FileOut);

            try
            {
                int b;
                long foundSeqOffset = -1;
                int searchByteCursor = 0;

                while ((b = sourceStream.ReadByte()) != -1)
                {
                    if (_sourceSeq[searchByteCursor] == b)
                    {
                        if (searchByteCursor == _sourceSeq.Length - 1)
                        {
                            targetStream.Write(_targetSeq, 0, _targetSeq.Length);
                            searchByteCursor = 0;
                            foundSeqOffset = -1;
                        }
                        else
                        {
                            if (searchByteCursor == 0)
                            {
                                foundSeqOffset = sourceStream.Position - 1;
                            }

                            ++searchByteCursor;
                        }
                    }
                    else
                    {
                        if (searchByteCursor == 0)
                        {
                            targetStream.WriteByte((byte)b);
                        }
                        else
                        {
                            targetStream.WriteByte(_sourceSeq[0]);
                            sourceStream.Position = foundSeqOffset + 1;
                            searchByteCursor = 0;
                            foundSeqOffset = -1;
                        }
                    }
                }
            }
            finally
            {
                sourceStream.Dispose();
                targetStream.Dispose();
            }
        }
    }
}