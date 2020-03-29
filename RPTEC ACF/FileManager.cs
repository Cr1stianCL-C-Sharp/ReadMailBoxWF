
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System;
using System.IO;
using System.Collections;

namespace RPTEC_ACF
{
    public class FileManager
    {

      
        public Boolean ScanFilePath(String path)
        {


            try
            {
           
                // Determine whether the directory exists.
                if (Directory.Exists(path))
                {
                    //Console.WriteLine("That path exists already.");
                    return true;
                }
                else
                {

                    //Try to create the directory.
                    DirectoryInfo di = Directory.CreateDirectory(path);
                    //Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));
                    return true;
                    // System.IO.Directory.CreateDirectory(path);                  
                }
            }
            catch (Exception e)
            {
                Logs.Log("FileManager.ScanFilePath", e.Message.ToString());
                return false;
            }


        }//termnino ScanFilePath


    }///termino clase
}///termino Namespace





    //try
    //        {
           
    //            // Determine whether the directory exists.
    //            if (Directory.Exists(path))
    //            {
    //                //Console.WriteLine("That path exists already.");
    //                return true;
    //            }
    //            else
    //            { 
    //            // Try to create the directory.
    //            DirectoryInfo di = Directory.CreateDirectory(path);
    //            //Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));
    //                return false;
    //            }

    //        }
    //        catch (Exception)
    //        {

    //            throw;
    //        }
//If you do not have at a minimum read-only permission to the directory, the Exists method will return false.
//The Exists method returns false if any error occurs while trying to determine if the specified file exists.
//This can occur in situations that raise exceptions such as passing a file name with invalid characters or too,
//many characters, a failing or missing disk, or if the caller does not have permission to read the file.
////////// For File.Exists, Directory.Exists
////////using System;
////////using System.IO;
////////using System.Collections;

////////public class RecursiveFileProcessor
////////{
////////    public static void Main(string[] args)
////////    {
////////        foreach (string path in args)
////////        {
////////            if (File.Exists(path))
////////            {
////////                // This path is a file
////////                ProcessFile(path);
////////            }
////////            else if (Directory.Exists(path))
////////            {
////////                // This path is a directory
////////                ProcessDirectory(path);
////////            }
////////            else
////////            {
////////                Console.WriteLine("{0} is not a valid file or directory.", path);
////////            }
////////        }
////////    }


////////    // Process all files in the directory passed in, recurse on any directories 
////////    // that are found, and process the files they contain.
////////    public static void ProcessDirectory(string targetDirectory)
////////    {
////////        // Process the list of files found in the directory.
////////        string[] fileEntries = Directory.GetFiles(targetDirectory);
////////        foreach (string fileName in fileEntries)
////////            ProcessFile(fileName);

////////        // Recurse into subdirectories of this directory.
////////        string[] subdirectoryEntries = Directory.GetDirectories(targetDirectory);
////////        foreach (string subdirectory in subdirectoryEntries)
////////            ProcessDirectory(subdirectory);
////////    }

////////    // Insert logic for processing found files here.
////////    public static void ProcessFile(string path)
////////    {
////////        Console.WriteLine("Processed file '{0}'.", path);
////////    }
////////}

