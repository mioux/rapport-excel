// Redirection Utility
// Author: Hai Vu (haivu2004 on Google mail)
using System;
using System.IO;

namespace Rapport
{    
    /// <summary>
    /// OutToFile is an easy way to redirect console output to a file.
    /// Usage:
    ///    Console.WriteLine("This text goes to the console by default");
    ///    using (OutToFile redir = new OutToFile("out.txt"))
    ///    {
    ///         Console.WriteLine("Contents of out.txt");
    ///    }
    ///    Console.WriteLine("This text goes to console again");
    ///
    /// </summary>

    public class OutToFile : IDisposable
    {
		StreamWriter fileOutput;
		private TextWriter oldOutput;
		TextWriter oldError;
		/// <summary>
		/// Create a new object to redirect the output
		/// </summary>
		/// <param name="outFileName">
		/// The name of the file to capture console output

		public OutToFile(string outFileName)
        {
            oldOutput = Console.Out;
            oldError = Console.Error;
            fileOutput = new StreamWriter(
                new FileStream(outFileName, FileMode.Append)
                );
            fileOutput.AutoFlush = true;
            Console.SetOut(fileOutput);
            Console.SetError(fileOutput);
        }

        // Dispose() is called automatically when the object    
        // goes out of scope    

        public void Dispose()
        {
            // Restore the console output
            Console.SetOut(oldOutput);
            Console.SetError(oldError);
            // Done with the file
            fileOutput.Close();
        }
    }
}