using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using OpenMcdf;

namespace MCDFWrapperTests
{
   [TestFixture]
    public class MCDFWrapperTests
    {
        protected const string testCVJFile = @"/media/dale/Data/dales_documents/projects/programming/my_source/libraries/c++/CVJParser/tests/TestResources/CVJFiles/solid8/2003.00.000.J01.WAL.x.cvj";
        protected const string testContentsFile = @"/media/dale/Data/dales_documents/projects/programming/my_source/libraries/c#/MCDFWrapper/Tests/MCDFWrapperTests/resources/contents.txt";

        [TestFixture]
        public class ConstructorMcdfWrapperTests : MCDFWrapperTests
        {
            [Test]
            public void ctor_ValidFiles_ConsecutiveConstructs()
            {
                string testCVJFileA =
                    @"/media/dale/Data/dales_documents/projects/programming/my_source/libraries/c++/CVJParser/tests/TestResources/CVJFiles/solid8/3214.01.026.J01.TEA.x.cvj";
                
                string testCVJFileB =
                    @"/media/dale/Data/dales_documents/projects/programming/my_source/libraries/c++/CVJParser/tests/TestResources/CVJFiles/solid8/3379.00.000.S03.PAN.x.cvj";
                
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFileA,
                    CFSUpdateMode.Update, CFSConfiguration.Default);
                
                mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFileB,
                    CFSUpdateMode.Update, CFSConfiguration.Default);
                
            }
        }

        [TestFixture]
        public class FunctionMcdfWrapperTests : MCDFWrapperTests
        {
            [Test]
            public void SetStreamData_ValidString_StreamDataSetAndCommitted()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                string streamName = "Contents";
                string streamData = "MCDFWRAPPERTEST";
                MCDFWrapper.EncodingType encodingType = MCDFWrapper.EncodingType.ASCII;

                mCDFWrapper.SetStreamData(streamName, streamData, encodingType);
                mCDFWrapper.Commit(true);
            }

            [Test]
            public void SetStreamData_LargeContents_StreamDataSetAndCommitted()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                string streamName = "Contents";
                string streamData = System.IO.File.ReadAllText(testContentsFile);

                long originalContentsLength = mCDFWrapper.CompoundFileSize;

                Assert.IsNotNull(streamData, nameof(streamData) + " is null");

                mCDFWrapper.SetStreamData(streamName, streamData, MCDFWrapper.EncodingType.UTF8);
                mCDFWrapper.Commit(true);
            }

            [Test]
            public void SetStreamByteData_TestString_StreamDataSet()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                string streamName = "Contents";
                string streamData = "Test String";

                mCDFWrapper.SetStreamByteData(streamName, Encoding.ASCII.GetBytes(streamData));
            }

            [Test]
            public void EmptyStreamData_ValidStreamName_StreamDataSetToEmptyAndComitted()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                mCDFWrapper.EmptyStream("Contents");
                mCDFWrapper.Commit(true);
            }

            [Test]
            public void GetListOfStreams_FileWithMoreThanOneStream_ReturnedListOfStreams()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                List<string> listOfStreams = mCDFWrapper.GetListOfStreams();
                Assert.IsTrue(listOfStreams.Count > 0);
            }

            [Test]
            public void GetArrayOfStreams_FileWithMoreThanOneStream_ReturnedArrayOfStreams()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                string[] arrayOfStreams = mCDFWrapper.GetArrayOfStreams();
                Assert.IsTrue(arrayOfStreams.Length > 0);
            }

            [Test]
            public void GetStreamContents_StreamWithContents_ReturnStreamContents()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                string streamContents = mCDFWrapper.GetStreamData("Contents",
                    MCDFWrapper.EncodingType.ASCII);

                Assert.IsTrue(!string.IsNullOrEmpty(streamContents));
            }
            
            [Test]
            public void GetStreamContentsBytes_StreamWithContents_ReturnStreamContentsBytes()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);

                byte[] streamContents = mCDFWrapper.GetStreamByteData("Contents");

                Assert.IsTrue(streamContents.Length > 0);
            }

            [Test]
            public void ShrinkCompoundFile_CompoundFile_FileShrunk()
            {
                MCDFWrapper.MCDFWrapper mCDFWrapper = new MCDFWrapper.MCDFWrapper(testCVJFile,
                    CFSUpdateMode.Update, CFSConfiguration.Default);
                
                mCDFWrapper.ShrinkCompundFile();  
            }
        }
    }
}