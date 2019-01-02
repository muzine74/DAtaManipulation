using System.Threading.Tasks;

namespace DataTransformation.Interface
{
    public interface IDataInterface
    {
        Task ReadXLSFile(string pathFile);
    }
}
