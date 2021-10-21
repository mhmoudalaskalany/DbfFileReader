using FluentScheduler;

namespace FileReader.Job
{
    public class DbfJob : IJob
    {
        public void Execute()
        {
            DbfReader.DbfReader.ReadUserInput();
        }
    }
}
