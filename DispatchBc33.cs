using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Collections;

using AIT.SM;
using AIT.ITF;
//using ExcelLibrary.SpreadSheet;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace AIT.XlsGen
{
    /// <summary>
    /// Generate File Diaptch BC33 XLS (excel 2003 format) 
    /// </summary>
    public class DispatchBc33
    {
        #region Public
        /// <summary>
        /// Generate excel files for Inbound and Outbound
        /// </summary>
        /// <param name="dataLog">Detail message data</param>
        /// <param name="dataSftp">Data SFTP setup which related to Data Log</param>
        /// <param name="dataConfig">Data Config of the message</param>
        /// <param name="fileConnectionString">Filename of main Connection String</param>
        public void Process(ref Log.DataLog dataLog, SftpServer.DataSftpServer dataSftp, MessageConfig.DataMessageConfig dataConfig, string fileConnectionString)
        {
            try
            {
                Config.DataConfig sysConfig = new Config().Get(Config.ID.BINPATH.ToString(), fileConnectionString);
                p_process(dataLog, dataSftp, sysConfig.valueString);
            }
            catch { throw; }
        }

        #endregion

        #region Private
        private void p_process(Log.DataLog dataLog, SftpServer.DataSftpServer dataSftp, string binPath)
        {
            try
            {
                string filename = dataLog.Filename;
                string partName = filename.Split('_')[2];
                string externOrderKey = partName.Split('.')[0];

                string destinationFile = string.Format("{0}{1}", binPath, filename);
                if (!File.Exists(destinationFile))
                {
                    IWorkbook workbook = new HSSFWorkbook();
                    ISheet sheet = workbook.CreateSheet("Dispatch");

                    int lastRow = 0;
                    foreach (Log.DataLogDetail detail in dataLog.LogDetail)
                    {
                        string[] cellValue = detail.LineMessage.Split(',');
                        IRow row = sheet.CreateRow(lastRow);

                        for (int i = 0; i < cellValue.Length; i++)
                        {
                            ICell cell = row.CreateCell(i);
                            cell.SetCellValue(cellValue[i].Replace("|", ",").ToString());
                        }
                        lastRow += 1;
                    }

                    ////merging cells
                    for (int i = 0; i < 8; i++) { sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(4, 5, i, i)); }
                    for (int i = 14; i < 18; i++) { sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(4, 5, i, i)); }

                    using (FileStream stream = new FileStream(destinationFile, FileMode.Create, FileAccess.Write)) { workbook.Write(stream); }
                }

                string finalDestinationFilename = dataSftp.LocalDirectoryArchieve + dataLog.Filename;
                if (File.Exists(finalDestinationFilename))
                    File.Delete(finalDestinationFilename);
                File.Move(destinationFile, finalDestinationFilename);
            }
            catch { throw; }
        }
        #endregion
    }
}
