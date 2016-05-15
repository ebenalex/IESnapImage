using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace EVFile
{
    public class FileInfo
    {
        #region Variables
        private string path;
        private string filename;
        private string extname;

        #endregion

        #region Fields

        public string Path
        {
            get
            {
                return path;
            }

            set
            {
                path = value;
            }
        }

        public string Filename
        {
            get
            {
                return filename;
            }


        }

        public string Filepath
        {
            get
            {
                return path + filename;
            }
        }

        public string Extname
        {
            get
            {
                return extname;
            }

            set
            {
                extname = value;
            }
        }

        #endregion

        #region constructions
        public FileInfo(string _path, string _extname)
        {
            path = _path;
            extname = _extname;
            genFileName();
        }
        #endregion

        #region private methods
        private void genFileName()
        {
            filename = Guid.NewGuid() + "." + extname;
            return;
        }
        #endregion

        #region public methods
        public void WriteText(string _contents)
        {
            System.IO.File.WriteAllText(path + filename, _contents);
        }

        public void MoveFile(string _filePath)
        {
            System.IO.File.Copy(_filePath, path + filename, true);
            System.IO.File.Delete(_filePath);
        }
        #endregion
    }
}
