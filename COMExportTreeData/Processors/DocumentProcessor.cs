using System;
using System.Runtime.InteropServices;
using COMExportTreeData.Models;
using INFITF;
using MECMOD;
using ProductStructureTypeLib;

namespace COMExportTreeData.Processors {
    /// <summary>
    /// 文档处理器 - 处理CATIA文档
    /// </summary>
    public static class DocumentProcessor {
        /// <summary>
        /// 处理单个文档
        /// </summary>
        public static NodeSchema ProcessDocument(Application catia, string docPath, string mode, int maxDepth) {
            Document doc = null;
            try {
                // 检查文件是否存在
                if (!System.IO.File.Exists(docPath)) {
                    Console.WriteLine($"文件不存在：{docPath}");
                    return null;
                }

                doc = catia.Documents.Open(docPath);
                Console.WriteLine($"打开文档：{docPath}");

                NodeSchema result = null;

                // 根据文档类型处理（产品/零件）
                string docName = doc.get_Name();
                if (docName.EndsWith(".CATProduct", StringComparison.OrdinalIgnoreCase)) {
                    ProductDocument productDoc = (ProductDocument)doc;
                    Product rootProduct = productDoc.Product;

                    // 处理产品结构树
                    result = ProductProcessor.ProcessProduct(rootProduct, mode, maxDepth, 0);
                }
                else if (docName.EndsWith(".CATPart", StringComparison.OrdinalIgnoreCase)) {
                    PartDocument partDoc = (PartDocument)doc;
                    Part part = partDoc.Part;

                    // 处理零件
                    result = PartProcessor.ProcessPart(part, mode, maxDepth, 0);
                }
                else {
                    Console.WriteLine($"不支持的文档类型：{docName}");
                }

                return result;
            }
            catch (Exception ex) {
                Console.WriteLine($"处理文档失败 [{docPath}]：{ex.Message}");
                return null;
            }
            finally {
                if (doc != null) {
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                }
            }
        }
    }
}