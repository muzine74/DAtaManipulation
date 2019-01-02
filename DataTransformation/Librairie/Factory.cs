using DataTransformation.Interface;

namespace DataTransformation.Librairie
{
    public class Factory
    {
        public IDataInterface GetObject(string type, string pathFile)
        {
            IDataInterface Folder;

            switch (type)
            {
                case ".xlsx":
                    Folder = new ExcelFolder(pathFile);
                    break;

                case ".txt":
                    Folder = new TextFolder();
                    break;
                default:
                    Folder = null;
                    break;
            }

            return Folder;
        }
    }
}
