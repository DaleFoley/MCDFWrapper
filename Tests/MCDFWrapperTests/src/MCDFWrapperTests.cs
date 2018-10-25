using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using OpenMcdf;


namespace MCDFWrapperTests
{
    [TestClass]
    public class MCDFWrapperTests
    {
        protected const string testCVJFile = @"E:\dales_documents\projects\programming\my_source\libraries\c++\CVJParser\tests\TestResources\CVJFiles\solid8\2003.00.000.J01.WAL.x.cvj";
        protected const string testContentsFile = @"E:\dales_documents\projects\programming\my_source\libraries\c#\MCDFWrapper\Tests\MCDFWrapperTests\resources\contents.txt";

        [TestClass]
        public class ConstructorTests : MCDFWrapperTests
        {
            
        }

        [TestClass]
        public class FunctionTests : MCDFWrapperTests
        {
            [TestMethod]
            public void SetStreamData_ValidString_StreamDataSetAndCommitted()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                string streamName = "Contents";
                string streamData = "MCDFWRAPPERTEST";
                MCDFWrapper.EncodingType encodingType = MCDFWrapper.EncodingType.ASCII;

                mCDFWrapper.SetStreamData(streamName, streamData, encodingType);
                mCDFWrapper.CompoundFile.Commit(true);
            }

            [TestMethod]
            public void SetStreamData_LargeContents_StreamDataSetAndCommitted()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                string streamName = "Contents";
                string streamData = System.IO.File.ReadAllText(testContentsFile);

                long originalContentsLength = mCDFWrapper.CompoundFileSize;

                Assert.IsNotNull(streamData, nameof(streamData) + " is null");

                mCDFWrapper.SetStreamData(streamName, streamData, MCDFWrapper.EncodingType.UTF8);
                mCDFWrapper.CompoundFile.Commit(true);
            }

            [TestMethod]
            public void EmptyStreamData_ValidStreamName_StreamDataSetToEmptyAndComitted()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                string streamName = "Contents";

                mCDFWrapper.EmptyStream(streamName);
                mCDFWrapper.CompoundFile.Commit(true);
            }
        }
    }
}
