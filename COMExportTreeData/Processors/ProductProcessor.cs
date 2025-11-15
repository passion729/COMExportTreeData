using System;
using COMExportTreeData.Models;
using ProductStructureTypeLib;

namespace COMExportTreeData.Processors {
    /// <summary>
    /// 产品处理器 - 处理Product文档
    /// </summary>
    public static class ProductProcessor {
        /// <summary>
        /// 处理产品结构树
        /// </summary>
        public static NodeSchema ProcessProduct(Product product, string mode, int maxDepth, int currentDepth) {
            try {
                var productNode = new NodeSchema {
                    Name = product.get_Name(),
                    Type = "Product"
                };

                // 添加产品属性（PartNumber和Nomenclature）
                string partNumber = product.get_PartNumber();
                string nomenclature = product.get_Nomenclature();

                if (!string.IsNullOrEmpty(partNumber)) {
                    productNode.AddProperty("PartNumber", partNumber);
                }

                if (!string.IsNullOrEmpty(nomenclature)) {
                    productNode.AddProperty("Nomenclature", nomenclature);
                }

                // 如果达到最大深度，不再递归
                if (maxDepth > 0 && currentDepth >= maxDepth) {
                    return productNode;
                }

                // 处理子产品
                Products childProducts = product.Products;
                for (int i = 1; i <= childProducts.Count; i++) {
                    try {
                        Product child = childProducts.Item(i);
                        NodeSchema childNode = ProcessProduct(child, mode, maxDepth, currentDepth + 1);
                        if (childNode != null) {
                            productNode.Children.Add(childNode);
                        }
                    }
                    catch (Exception ex) {
                        Console.WriteLine($"处理子产品失败：{ex.Message}");
                    }
                }

                return productNode;
            }
            catch (Exception ex) {
                Console.WriteLine($"处理产品失败：{ex.Message}");
                return null;
            }
        }
    }
}