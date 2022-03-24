using System.ServiceProcess;

namespace EDI_FTP
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new FTPService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
