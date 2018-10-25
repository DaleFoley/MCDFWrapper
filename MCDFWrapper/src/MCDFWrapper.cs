using System.Text;

using OpenMcdf;

namespace MCDFWrapper
{
    public enum EncodingType
    {
        ASCII,
        UTF7,
        UTF8,
        Unicode,
        UTF32,
        BigEndianUnicode
    }

    public class MCDFWrapper
    {
        public CompoundFile CompoundFile;

        private readonly string _pathOfCompoundFile;
        private readonly CFSUpdateMode _currentUpdateMode;
        private readonly CFSConfiguration _currentConfigParameters;

        public long CompoundFileSize => this.CompoundFile.RootStorage.Size;

        /// <summary>
        /// Load an existing compound file.
        /// </summary>
        /// <param name="fileName">Compound file to read from</param>
        /// <param name="sectorRecycle">If true, recycle unused sectors</param>
        /// <param name="updateMode">Select the update mode of the underlying data file</param>
        public MCDFWrapper(string fileName, CFSUpdateMode updateMode, CFSConfiguration configParameters)
        {
            this.CompoundFile = new CompoundFile(fileName, updateMode, configParameters);

            this._pathOfCompoundFile = fileName;
            this._currentUpdateMode = updateMode;
            this._currentConfigParameters = configParameters;
        }

        /// <summary>
        /// Remove all existing stream data and write new stream data.
        /// </summary>
        /// <param name="streamName">The name of the stream to operate on.</param>
        /// <param name="content">The content to be written to stream.</param>
        /// <param name="encoding">The encoding to be used.</param>
        public void SetStreamData(string streamName, string content, EncodingType encoding)
        {
            byte[] dataToBeSet;

            switch (encoding)
            {
                case EncodingType.ASCII:
                    dataToBeSet = Encoding.ASCII.GetBytes(content);
                    break;
                case EncodingType.BigEndianUnicode:
                    dataToBeSet = Encoding.BigEndianUnicode.GetBytes(content);
                    break;
                case EncodingType.Unicode:
                    dataToBeSet = Encoding.Unicode.GetBytes(content);
                    break;
                case EncodingType.UTF32:
                    dataToBeSet = Encoding.UTF32.GetBytes(content);
                    break;
                case EncodingType.UTF8:
                    dataToBeSet = Encoding.UTF8.GetBytes(content);
                    break;
                case EncodingType.UTF7:
                    dataToBeSet = Encoding.UTF7.GetBytes(content);
                    break;
                default:
                    dataToBeSet = Encoding.ASCII.GetBytes(content);
                    break;
            }

            CFStream streamToSetData = this.CompoundFile.RootStorage.GetStream(streamName);
            streamToSetData.SetData(dataToBeSet);
        }

        /// <summary>
        /// Empties the stream, commits the changes to disk and closes the file, then shrinks the compound file
        /// and re-opens the compound file adding the original stream back in.
        /// </summary>
        /// <param name="streamName">The stream to empty.</param>
        public void EmptyStream(string streamName)
        {
            this.CompoundFile.RootStorage.Delete(streamName);
            this.CompoundFile.Commit(true);

            this.CompoundFile.Close();

            CompoundFile.ShrinkCompoundFile(this._pathOfCompoundFile);

            this.CompoundFile = new CompoundFile(_pathOfCompoundFile, _currentUpdateMode, _currentConfigParameters);
            this.CompoundFile.RootStorage.AddStream(streamName);
        }
    }
}
