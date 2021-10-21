using System;
using System.Threading;
using FileReader.Job;
using FluentScheduler;


namespace FileReader
{
    /// <summary>
    /// Kick Off
    /// </summary>
    public class Program
    {
        static ManualResetEvent _quitEvent = new ManualResetEvent(false);

        #region Public Methods

        /// <summary>
        /// Kick Off
        /// </summary>
        /// <param name="args"></param>
        public static void Main(string[] args)
        {
           
            try
            {
                Console.CancelKeyPress += (sender, eArgs) => {
                    _quitEvent.Set();
                    eArgs.Cancel = true;
                };
                JobManager.Initialize(new DbfRegistry());
                // kick off asynchronous stuff 

                _quitEvent.WaitOne();

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }


            //ReadUserInput();
            //Console.ForegroundColor = ConsoleColor.DarkYellow;
        }

        #endregion


        
    }
}