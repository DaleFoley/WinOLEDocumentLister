using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenMcdf;
using System.Diagnostics;

namespace WinOLEDocumentLister
{
    class WinOLEDocumentLister
    {
        private static MCDFWrapper.MCDFWrapper _mcdfWrapper;
        private static Stopwatch _stopwatch;

        private static bool _isWriteFilesToSeparateDirectoriesBasedOnEncoding;

        private const string rootEntry = "Root Entry";
        private const string defaultSearchPattern = "*";

        static void Main(string[] args)
        {
            try
            {
                int argumentLength = args.Length;
                if (argumentLength < 1)
                {
                    Console.WriteLine("Expected argument for source file/directory to process. Got [" + argumentLength + "] number of arguments.");
                    Console.ReadKey();

                    return;
                }

                string pathFileOrDirectoryToBeProcessed = args[0];

                string searchPatternForDirectory = null;
                if (argumentLength >= 2)
                {
                    searchPatternForDirectory = args[1];
                }

                if(argumentLength >= 3)
                {                    
                    Boolean.TryParse(args[2], out _isWriteFilesToSeparateDirectoriesBasedOnEncoding);
                }

                Console.WriteLine("Starting to process OLE documents.");
                Console.WriteLine(Environment.NewLine);

                _stopwatch = new Stopwatch();
                _stopwatch.Start();

                bool isDirectory = Directory.Exists(pathFileOrDirectoryToBeProcessed);
                if(isDirectory)
                {
                    string[] filesToBeProcessed = null;
                    if (searchPatternForDirectory != null)
                    {
                        Console.WriteLine("Processing directory. Will get files based on search pattern [" + searchPatternForDirectory + "]");
                        filesToBeProcessed = Directory.GetFiles(pathFileOrDirectoryToBeProcessed, searchPatternForDirectory);
                    }
                    else
                    {
                        Console.WriteLine("Processing directory. Will get files based on default search pattern [" + defaultSearchPattern + "]");
                        filesToBeProcessed = Directory.GetFiles(pathFileOrDirectoryToBeProcessed, defaultSearchPattern);
                    }

                    foreach (string fileBeingProcessed in filesToBeProcessed)
                    {
                        string pathDirectoryToSaveTo = pathFileOrDirectoryToBeProcessed +
                                                        Path.DirectorySeparatorChar +
                                                        Path.GetFileNameWithoutExtension(fileBeingProcessed);
                        Directory.CreateDirectory(pathDirectoryToSaveTo);

                        _mcdfWrapper = new MCDFWrapper.MCDFWrapper(fileBeingProcessed, CFSUpdateMode.ReadOnly, CFSConfiguration.Default);
                        List<string> streamsInFile = _mcdfWrapper.GetListOfStreams(true);

                        Console.WriteLine("------Begin Stream Extraction On File [" + fileBeingProcessed + "]------");
                        WriteStreamsToFile(streamsInFile, pathDirectoryToSaveTo);
                        Console.WriteLine("------Finished Stream Extraction On File [" + fileBeingProcessed + "]------");
                    }
                }
                else
                {
                    string pathDirectoryToSaveTo = Path.GetDirectoryName(pathFileOrDirectoryToBeProcessed) +
                                                    Path.DirectorySeparatorChar +
                                                    Path.GetFileNameWithoutExtension(pathFileOrDirectoryToBeProcessed);
                    Directory.CreateDirectory(pathDirectoryToSaveTo);

                    _mcdfWrapper = new MCDFWrapper.MCDFWrapper(pathFileOrDirectoryToBeProcessed, CFSUpdateMode.ReadOnly, CFSConfiguration.Default);
                    List<string> streamsInFile = _mcdfWrapper.GetListOfStreams(true);

                    Console.WriteLine("------Begin Stream Extraction On File [" + pathFileOrDirectoryToBeProcessed + "]------");
                    WriteStreamsToFile(streamsInFile, pathDirectoryToSaveTo);
                    Console.WriteLine("------Finished Stream Extraction On File [" + pathFileOrDirectoryToBeProcessed + "]------");
                }
            }
            catch(Exception ex)
            {
                string err = "Encountered error [" + ex.Message + "]" +
                            Environment.NewLine +
                            Environment.NewLine +
                            "Stack Trace [" + ex.StackTrace + "]";
                Console.WriteLine(err);

                Console.ReadKey();
            }

            _stopwatch.Stop();

            Console.WriteLine("Finished processing OLE documents. Time elapsed hours [" +
                            _stopwatch.Elapsed.Hours + "] minutes [" +
                            _stopwatch.Elapsed.Minutes + "] seconds [" +
                            _stopwatch.Elapsed.Seconds + "]");
            Console.ReadKey();
        }

        private static void WriteStreamsToFile(List<string> streams, string pathSaveDiretory)
        {
            foreach (string stream in streams)
            {
                //Ignore root entry.
                if (stream.Equals(rootEntry)) { continue; }

                string pathStreamFile = pathSaveDiretory + Path.DirectorySeparatorChar + stream + ".txt";
                Console.WriteLine("------Begin Stream Write------ [" + stream + "]");
                Console.WriteLine("Writing stream [" + stream + "] to file [" + pathStreamFile + "]");

                byte[] streamDataByte = _mcdfWrapper.GetStreamByteData(stream);
                File.WriteAllBytes(pathStreamFile, streamDataByte);

                Console.WriteLine("Finished writing stream [" + stream + "] to file [" + pathStreamFile + "]");
                Console.WriteLine("------End Stream Write------ [" + stream + "]");
            }
        }
    }
}
