using System;
using System.IO;
using System.Text;


namespace UploadAndMapingNew
{
   sealed class WriteToLog
    {
        string writePath;

        public WriteToLog(string writePath)
        {
            this.writePath = writePath;
        }

        public void writeToLog(string messageText)
        {
            using (StreamWriter sw = new StreamWriter(writePath, true, Encoding.Default))
            {
                DateTime currentTime = DateTime.UtcNow;
                sw.WriteLine(messageText+ currentTime.ToString());
                sw.Dispose();

            }
        }
    }
}
