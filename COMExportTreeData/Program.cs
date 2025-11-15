using System;
using System.Collections.Generic;
using System.Text;
using COMExportTreeData.Helpers;
using COMExportTreeData.Models;
using COMExportTreeData.Processors;
using INFITF;
using Newtonsoft.Json;

namespace COMExportTreeData {
    static class Program {
        static void Main(string[] args) {
            // 校验参数（C#中args不包含程序名，直接从args[0]开始）
            if (args.Length != 4) {
                Console.WriteLine("参数错误！格式：程序名 [零件路径] [Json路径] [搜索代数] [模式(all/flatten)]");
                return;
            }

            string docPath = args[0]; // 零件路径
            string jsonPath = args[1]; // Json路径
            string mode = args[3]; // 模式

            // 解析搜索代数参数
            if (!int.TryParse(args[2], out var genNum) || genNum < 0) {
                Console.WriteLine("搜索代数参数错误！必须是非负整数");
                return;
            }

            // 连接CATIA（若未运行则启动）
            Application catia;
            try {
                catia = CatiaConnectionHelper.ConnectToCatia();
            }
            catch (Exception ex) {
                Console.WriteLine($"连接CATIA失败：{ex.Message}");
                return;
            }

            try {
                // 解析文档路径列表（支持分号分隔的多个路径）
                string[] docPaths = docPath.Split(';');
                if (docPaths.Length == 0) {
                    Console.WriteLine("未找到文档路径");
                    return;
                }

                // 准备JSON输出结构
                var resultList = new List<NodeSchema>();

                // 循环处理每个文档
                foreach (string path in docPaths) {
                    if (string.IsNullOrWhiteSpace(path))
                        continue;

                    var docResult = DocumentProcessor.ProcessDocument(catia, path.Trim(), mode, genNum);
                    if (docResult != null) {
                        resultList.Add(docResult);
                    }
                }

                // 将结果写入JSON文件
                if (resultList.Count > 0) {
                    var json =
                        // 如果只有一个文档，直接输出对象；否则输出数组
                        resultList.Count == 1
                            ? JsonConvert.SerializeObject(resultList[0], Formatting.Indented)
                            : JsonConvert.SerializeObject(resultList, Formatting.Indented);

                    System.IO.File.WriteAllText(jsonPath, json, new UTF8Encoding(false));
                    Console.WriteLine($"成功导出到：{jsonPath}");
                }
                else {
                    Console.WriteLine("未生成任何数据");
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"处理失败：{ex.Message}");
                Console.WriteLine($"堆栈跟踪：{ex.StackTrace}");
            }
            finally {
                // 释放资源
                CatiaConnectionHelper.ReleaseCatia(catia);
            }
        }
    }
}