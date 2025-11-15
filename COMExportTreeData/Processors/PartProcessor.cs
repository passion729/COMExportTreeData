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
                    return ProcessPartFull(part, part, maxDepth, 0);
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
        private static NodeSchema ProcessPartFull(Part rootPart, object iObj, int maxDepth, int currentDepth) {
            if (iObj == null) {
                return null;
            }

            try {
                // 1. 获取节点名称和类型
                string nodeName = CatiaObjectHelper.GetObjectName(iObj);
                string nodeType = CatiaObjectHelper.GetObjectType(iObj);

                // 2. 如果是参数节点，只返回 name 和 value
                if (iObj is Parameter parameter) {
                    string paramValue = ParameterHelper.GetParameterValue(parameter);
                    return new NodeSchema {
                        Name = nodeName,
                        Value = paramValue,
                        Children = null  // 参数节点不需要 children
                    };
                }

                // 3. 非参数节点，正常处理
                var node = new NodeSchema {
                    Name = nodeName,
                    Type = nodeType
                };

                // 4. 检查是否达到最大深度
                if (maxDepth > 0 && currentDepth >= maxDepth) {
                    return node;
                }

                // 5. 递归处理子节点（现在包括参数节点）
                List<object> children = CatiaObjectHelper.GetChildren(rootPart, iObj);
                if (children != null && children.Count > 0) {
                    foreach (var child in children) {
                        NodeSchema childNode = ProcessPartFull(rootPart, child, maxDepth, currentDepth + 1);
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
        /// 扁平化模式：展平所有嵌套节点到一层，使用路径作为名称
        /// </summary>
        private static NodeSchema ProcessPartFlatten(object iObj, int maxDepth, int currentDepth) {
            if (iObj == null) {
                return null;
            }

            try {
                Part part = iObj as Part;
                if (part == null) {
                    return null;
                }

                // 创建根节点
                var rootNode = new NodeSchema {
                    Name = CatiaObjectHelper.GetObjectName(part),
                    Type = "Part"
                };

                // 收集所有扁平化的节点，跳过 Part 本身，从第一层子节点开始
                List<NodeSchema> flattenedNodes = new List<NodeSchema>();
                List<object> firstLevelChildren = CatiaObjectHelper.GetChildren(part, part);
                
                if (firstLevelChildren != null) {
                    foreach (var child in firstLevelChildren) {
                        if (!(child is Parameter)) {
                            // 从第一层子节点开始，路径前缀为空
                            FlattenNode(part, child, "", flattenedNodes, maxDepth, 1);
                        }
                    }
                }

                // 将扁平化的节点添加到根节点
                rootNode.Children.AddRange(flattenedNodes);

                return rootNode;
            }
            catch (Exception ex) {
                Console.WriteLine($"ProcessPartFlatten异常：{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 递归扁平化节点
        /// </summary>
        private static void FlattenNode(Part rootPart, object obj, string pathPrefix, List<NodeSchema> flattenedNodes, int maxDepth, int currentDepth) {
            if (obj == null) return;

            // 检查最大深度
            if (maxDepth > 0 && currentDepth >= maxDepth) {
                return;
            }

            string nodeName = CatiaObjectHelper.GetObjectName(obj);
            string nodeType = CatiaObjectHelper.GetObjectType(obj);
            
            // 构建当前路径
            string currentPath = string.IsNullOrEmpty(pathPrefix) ? nodeName : $"{pathPrefix}/{nodeName}";

            // 获取子节点
            List<object> children = CatiaObjectHelper.GetChildren(rootPart, obj);
            
            // 分离结构子节点和参数子节点
            List<object> structureChildren = new List<object>();
            List<Parameter> parameterChildren = new List<Parameter>();
            
            if (children != null) {
                foreach (var child in children) {
                    if (child is Parameter param) {
                        parameterChildren.Add(param);
                    }
                    else {
                        structureChildren.Add(child);
                    }
                }
            }

            // 如果有参数子节点，创建一个扁平化节点
            if (parameterChildren.Count > 0) {
                var flatNode = new NodeSchema {
                    Name = currentPath,
                    Type = nodeType
                };

                // 添加参数作为子节点
                foreach (var param in parameterChildren) {
                    string paramName = CatiaObjectHelper.GetObjectName(param);
                    string paramValue = ParameterHelper.GetParameterValue(param);
                    
                    flatNode.Children.Add(new NodeSchema {
                        Name = paramName,
                        Value = paramValue,
                        Children = null
                    });
                }

                flattenedNodes.Add(flatNode);
            }

            // 递归处理结构子节点
            foreach (var child in structureChildren) {
                FlattenNode(rootPart, child, currentPath, flattenedNodes, maxDepth, currentDepth + 1);
            }
        }
    }
}