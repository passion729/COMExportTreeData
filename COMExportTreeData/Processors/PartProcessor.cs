using System;
using System.Collections.Generic;
using COMExportTreeData.Helpers;
using COMExportTreeData.Models;
using KnowledgewareTypeLib;
using MECMOD;

namespace COMExportTreeData.Processors {
    /// <summary>
    /// 零件处理器 - 处理Part文档
    /// </summary>
    public static class PartProcessor {
        /// <summary>
        /// 处理单个零件（生成JSON）
        /// </summary>
        public static NodeSchema ProcessPart(Part part, string mode, int maxDepth, int currentDepth) {
            try {
                if (mode == "all") {
                    return ProcessPartFull(part, maxDepth, 0);
                }
                else if (mode == "flatten") {
                    return ProcessPartFlatten(part, maxDepth, 0);
                }
                else {
                    Console.WriteLine($"不支持的模式：{mode}，请使用 'all' 或 'flatten'");
                    return null;
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"处理零件失败：{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 完整模式：递归遍历特征层级结构
        /// </summary>
        private static NodeSchema ProcessPartFull(object iObj, int maxDepth, int currentDepth) {
            if (iObj == null) {
                return null;
            }

            try {
                var node = new NodeSchema();

                // 1. 获取节点名称和类型
                string nodeName = CatiaObjectHelper.GetObjectName(iObj);
                string nodeType = CatiaObjectHelper.GetObjectType(iObj);

                node.Name = nodeName;
                node.Type = nodeType;

                // 2. 读取节点下的所有参数作为属性
                if (iObj is Parameter parameter) {
                    // 如果节点本身是参数，添加其值
                    string paramValue = ParameterHelper.GetParameterValue(parameter);
                    if (!string.IsNullOrEmpty(paramValue)) {
                        node.AddProperty(nodeName, paramValue);
                    }
                }
                else {
                    // 如果节点不是参数，尝试获取它的所有参数
                    try {
                        List<Parameter> nodeParams = ParameterHelper.GetDirectParameters(iObj);
                        foreach (var param in nodeParams) {
                            try {
                                string pName = param.get_Name();
                                // 简化参数名（去掉路径前缀）
                                int lastSlash = pName.LastIndexOf("\\");
                                if (lastSlash >= 0 && lastSlash < pName.Length - 1) {
                                    pName = pName.Substring(lastSlash + 1);
                                }

                                string pValue = ParameterHelper.GetParameterValue(param);
                                if (!string.IsNullOrEmpty(pName)) {
                                    node.AddProperty(pName, pValue);
                                }
                            }
                            catch (Exception ex) {
                                Console.WriteLine($"读取参数失败：{ex.Message}");
                            }
                        }
                    }
                    catch {
                    }
                }

                // 3. 检查是否达到最大深度
                if (maxDepth > 0 && currentDepth >= maxDepth) {
                    return node;
                }

                // 4. 递归处理子节点
                List<object> children = CatiaObjectHelper.GetChildren(iObj);
                if (children != null && children.Count > 0) {
                    foreach (var child in children) {
                        NodeSchema childNode = ProcessPartFull(child, maxDepth, currentDepth + 1);
                        if (childNode != null) {
                            node.Children.Add(childNode);
                        }
                    }
                }

                return node;
            }
            catch (Exception ex) {
                Console.WriteLine($"ProcessPartFull异常：{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 扁平化模式：解析特征参数
        /// </summary>
        private static NodeSchema ProcessPartFlatten(object iObj, int maxDepth, int currentDepth) {
            if (iObj == null) {
                return null;
            }

            try {
                // 1. 获取零件的参考平面
                Part part = iObj as Part;
                if (part == null) {
                    return null;
                }

                object spXY = null, spYZ = null, spZX = null;

                try {
                    OriginElements origin = part.OriginElements;
                    spXY = origin.PlaneXY;
                    spYZ = origin.PlaneYZ;
                    spZX = origin.PlaneZX;
                }
                catch (Exception ex) {
                    Console.WriteLine($"获取参考平面失败：{ex.Message}");
                }

                // 2. 创建根节点
                var node = new NodeSchema {
                    Name = CatiaObjectHelper.GetObjectName(part),
                    Type = "Part"
                };

                // 3. 检查是否达到最大深度
                if (maxDepth > 0 && currentDepth >= maxDepth) {
                    return node;
                }

                // 4. 遍历子节点
                List<object> children = CatiaObjectHelper.GetChildren(iObj);
                if (children != null && children.Count > 0) {
                    foreach (var child in children) {
                        // 过滤参考平面和机械工具对象
                        if (spXY != null && ReferenceEquals(child, spXY)) continue;
                        if (spYZ != null && ReferenceEquals(child, spYZ)) continue;
                        if (spZX != null && ReferenceEquals(child, spZX)) continue;
                        if (CatiaObjectHelper.IsMechanicalTool(child)) continue;

                        try {
                            // 创建子节点
                            var childNode = new NodeSchema {
                                Name = CatiaObjectHelper.GetObjectName(child),
                                Type = CatiaObjectHelper.GetObjectType(child)
                            };

                            // 提取子节点下的所有参数作为属性
                            List<Parameter> allParams = ParameterHelper.GetAllParameters(child);
                            foreach (var param in allParams) {
                                try {
                                    string paramName = param.get_Name().Replace("\\", "/");
                                    int lastSlashIndex = paramName.LastIndexOf("/");
                                    if (lastSlashIndex >= 0 && lastSlashIndex < paramName.Length - 1) {
                                        paramName = paramName.Substring(lastSlashIndex + 1);
                                    }

                                    string paramValue = ParameterHelper.GetParameterValue(param);

                                    // 将参数作为属性添加到子节点，而不是作为子节点
                                    if (!string.IsNullOrEmpty(paramName)) {
                                        childNode.AddProperty(paramName, paramValue);
                                    }
                                }
                                catch (Exception ex) {
                                    Console.WriteLine($"处理参数失败：{ex.Message}");
                                }
                            }

                            node.Children.Add(childNode);
                        }
                        catch (Exception ex) {
                            Console.WriteLine($"处理子节点失败：{ex.Message}");
                        }
                    }
                }

                return node;
            }
            catch (Exception ex) {
                Console.WriteLine($"ProcessPartFlatten异常：{ex.Message}");
                return null;
            }
        }
    }
}