using System.Collections.Generic;
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
        private CompoundFile _compoundFile;

        private readonly string _pathOfCompoundFile;
        private readonly CFSUpdateMode _currentUpdateMode;
        private readonly CFSConfiguration _currentConfigParameters;

        public long CompoundFileSize => this._compoundFile.RootStorage.Size;

        /// <summary>
        /// Load an existing compound file.
        /// </summary>
        /// <param name="fileName">Compound file to read from</param>
        /// <param name="sectorRecycle">If true, recycle unused sectors</param>
        /// <param name="updateMode">Select the update mode of the underlying data file</param>
        public MCDFWrapper(string fileName, CFSUpdateMode updateMode, CFSConfiguration configParameters)
        {
            this._compoundFile = new CompoundFile(fileName, updateMode, configParameters);

            this._pathOfCompoundFile = fileName;
            this._currentUpdateMode = updateMode;
            this._currentConfigParameters = configParameters;
        }

        /// <summary>
        /// Return a list of all streams in this compound file.
        /// </summary>
        /// <returns>string list representing all directory names.</returns>
        public List<string> GetListOfStreams()
        {
            List<string> rtn = new List<string>();

            int numberOfDirectories = this._compoundFile.GetNumDirectories();
            for(var i = 0; i < numberOfDirectories; i++)
            {
                rtn.Add(this._compoundFile.GetNameDirEntry(i));
            }

            return rtn;
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

            CFStream streamToSetData = this._compoundFile.RootStorage.GetStream(streamName);
            streamToSetData.SetData(dataToBeSet);
        }

        /// <summary>
        /// Return string contents of stream in compound file.
        /// </summary>
        /// <param name="streamName">The name of the stream to operate on.</param>
        /// <param name="encoding">The encoding to be used.</param>
        /// <returns>string contents of stream.</returns>
        public string GetStreamData(string streamName, EncodingType encoding)
        {
            string rtn = null;

            CFStream streamToGetDataFrom = this._compoundFile.RootStorage.GetStream(streamName);
            byte[] streamData = streamToGetDataFrom.GetData();

            switch (encoding)
            {
                case EncodingType.ASCII:
                    rtn = Encoding.ASCII.GetString(streamData);
                    break;
                case EncodingType.BigEndianUnicode:
                    rtn = Encoding.BigEndianUnicode.GetString(streamData);
                    break;
                case EncodingType.Unicode:
                    rtn = Encoding.Unicode.GetString(streamData);
                    break;
                case EncodingType.UTF32:
                    rtn = Encoding.UTF32.GetString(streamData);
                    break;
                case EncodingType.UTF8:
                    rtn = Encoding.UTF8.GetString(streamData);
                    break;
                case EncodingType.UTF7:
                    rtn = Encoding.UTF7.GetString(streamData);
                    break;
                default:
                    rtn = Encoding.ASCII.GetString(streamData);
                    break;
            }

            return rtn;
        }

        /// <summary>
        /// Empties the stream, commits the changes to disk and closes the file, then shrinks the compound file
        /// and re-opens the compound file adding the original stream back in.
        /// </summary>
        /// <param name="streamName">The stream to empty.</param>
        public void EmptyStream(string streamName)
        {
            this._compoundFile.RootStorage.Delete(streamName);
            this._compoundFile.Commit(true);

            this._compoundFile.Close();

            CompoundFile.ShrinkCompoundFile(this._pathOfCompoundFile);

            this._compoundFile = new CompoundFile(_pathOfCompoundFile, _currentUpdateMode, _currentConfigParameters);
            this._compoundFile.RootStorage.AddStream(streamName);
        }

        /// <summary>
        /// Commit data changes since the previously commit operation
        /// to the underlying supporting stream or file on the disk.
        /// </summary>
        /// <param name="releaseMemory">If true, release loaded sectors to limit memory usage but reduces following read operations performance</param>
        /// <remarks>
        /// This method can be used only if 
        /// the supporting stream has been opened in 
        /// </remarks>
        public void Commit(bool releaseMemory)
        {
            this._compoundFile.Commit(releaseMemory);
        }
    }
}
