using BulkRegistrationProcesses.Models;
using EFCore.BulkExtensions;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace BulkRegistrationProcesses
{
    public class Worker : BackgroundService
    {
        private static Context dbContext;
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }
 
        //Sql Db
        private void SqlLogin()
        {
            _logger.LogInformation("Sql Login Entry {time}", DateTimeOffset.Now);
            string projectPath = AppDomain.CurrentDomain.BaseDirectory.Split(new String[] { @"bin\" }, StringSplitOptions.None)[0];
            IConfigurationRoot _configuration = new ConfigurationBuilder()
     .SetBasePath(projectPath)
     .AddJsonFile("appsettings.json")
     .Build();
            string constr = _configuration.GetConnectionString("MSSQLConnection");
            var contextOptions = new DbContextOptionsBuilder<DbContext>()
               .UseSqlServer(constr)
               .Options;
            dbContext = new Context(contextOptions);
            _logger.LogInformation("Sql Login Success Time: {time}", DateTimeOffset.Now);
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
        backToTap:
            try
            {
                _logger.LogInformation("Service Start Time: {time}", DateTimeOffset.Now);

                SqlLogin();

                while (!stoppingToken.IsCancellationRequested)
                {
                    _logger.LogInformation("While Beginning Time: {time}", DateTimeOffset.Now);
                    await ExcelRegister();
                    _logger.LogInformation("While End Time: {time}", DateTimeOffset.Now);

                    await Task.Delay((int)TimeSpan.FromSeconds(5).TotalMilliseconds, stoppingToken);
                }
            }
            catch (Exception ex)
            {
                WriteToFile(DateTime.Now.ToString() + "\n" + ex.Message);
                _logger.LogError(ex.Message + " Time: {time}", DateTimeOffset.Now);
            }
            finally
            {
                _logger.LogInformation("Service Stop Time: {time}", DateTimeOffset.Now);
            }
            goto backToTap;
        }

        private async Task ExcelRegister()
        {
            try
            {
                #region List
                List<Stock> stocks = new List<Stock>();
                #endregion

                var fullPath = Directory.GetCurrentDirectory();
                var file = fullPath + "/Stock.xlsx";

                _logger.LogInformation(file + " File Name Start Time: {time}", DateTimeOffset.Now);

                if (File.Exists(file) && file.Length > 0)
                {
                    _logger.LogInformation("File Open Time: {time}", DateTimeOffset.Now);

                    var fileName = file;

                    List<string> rowList = new List<string>();
                    ISheet sheet;
                    using (var stream = new FileStream(fileName, FileMode.Open))
                    {
                        stream.Position = 0;
                        XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);
                        sheet = xssWorkbook.GetSheetAt(0);
                        IRow headerRow = sheet.GetRow(0);

                        int cellCount = headerRow.LastCellNum;

                        for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null) continue;
                            if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                            for (int j = row.FirstCellNum; j < cellCount; j++)
                            {
                                if (row.GetCell(j) != null)
                                    if (!string.IsNullOrEmpty(row.GetCell(j).ToString()) && !string.IsNullOrWhiteSpace(row.GetCell(j).ToString()))
                                        rowList.Add(row.GetCell(j).ToString());
                            }
                            if (rowList.Count > 0)
                            {
                                _logger.LogInformation("Excel Read Start Row Number :" + i + " Time: {time}", DateTimeOffset.Now);

                                Stock stock = new Stock();
                                stock.Name = rowList[0];
                                stocks.Add(stock);

                                rowList.Clear();

                                _logger.LogInformation("Excel Read Finish Time: {time}", DateTimeOffset.Now);
                            }
                        }

                        using (var transaction = dbContext.Database.BeginTransaction())
                        {
                            _logger.LogInformation("Bulk Insert Start Time: {time}", DateTimeOffset.Now);

                            if (stocks.Count() != 0)
                                //Insert list data using BulkInsert
                                await dbContext.BulkInsertOrUpdateAsync(stocks);
         

                            //Commit, save changes
                            transaction.Commit();

                            _logger.LogInformation("Bulk Insert Finish Time: {time}", DateTimeOffset.Now);
                        }

                        File.Delete(file);
                        _logger.LogInformation("File Delete Time: {time}", DateTimeOffset.Now);
                        _logger.LogInformation("Register Success Time: {time}", DateTimeOffset.Now);
                    }
                }
                else
                {
                    _logger.LogInformation("No records found to be processed Time: {time}", DateTimeOffset.Now);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void WriteToFile(string Message)
        {
            #region Log Write
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            #endregion
        }
    }
}
