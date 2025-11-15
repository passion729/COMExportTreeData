using System;
using System.Collections.Generic;
using INFITF;
using KnowledgewareTypeLib;
using MECMOD;
using ProductStructureTypeLib;

namespace COMExportTreeData.Helpers {
    /// <summary>
    /// CATIA对象辅助类 - 提供通用的对象操作方法
    /// </summary>
    public static class CatiaObjectHelper {
        /// <summary>
        /// 获取对象名称
        /// </summary>
        public static string GetObjectName(object obj) {
            if (obj == null) return "Unknown";

            try {
                if (obj is Part part) return part.get_Name();
                if (obj is HybridBody hb) return hb.get_Name();
                if (obj is Body body) return body.get_Name();
                if (obj is HybridShape hs) return hs.get_Name();
                if (obj is Shape shape) return shape.get_Name();
                if (obj is Parameter param) return param.get_Name();

                // 尝试反射获取Name属性
                var nameProperty = obj.GetType().GetProperty("Name");
                if (nameProperty != null) {
                    return nameProperty.GetValue(obj)?.ToString() ?? "Unknown";
                }

                return obj.GetType().Name;
            }
            catch {
                return "Unknown";
            }
        }

        /// <summary>
        /// 获取对象类型
        /// </summary>
        public static string GetObjectType(object obj) {
            if (obj == null) return "Unknown";

            try {
                if (obj is Part) return "Part";
                if (obj is HybridBody) return "HybridBody";
                if (obj is Body) return "Body";
                if (obj is HybridShape) return "HybridShape";
                if (obj is Shape) return "Shape";
                if (obj is Parameter) return "Parameter";
                if (obj is Product) return "Product";

                return obj.GetType().Name;
            }
            catch {
                return "Unknown";
            }
        }

        /// <summary>
        /// 判断对象是否为机械工具
        /// </summary>
        public static bool IsMechanicalTool(object obj) {
            if (obj == null) return false;

            try {
                string typeName = obj.GetType().Name;
                return typeName.Contains("MechanicalTool") || typeName == "OriginElements";
            }
            catch {
                return false;
            }
        }

        /// <summary>
        /// 获取对象的子节点集合（包括结构子节点和参数子节点）
        /// </summary>
        public static List<object> GetChildren(Part rootPart, object obj) {
            List<object> children = new List<object>();

            if (obj == null) return children;

            try {
                // Part的子节点：几何集和实体Body
                if (obj is Part part) {
                    HybridBodies hybridBodies = part.HybridBodies;
                    for (int i = 1; i <= hybridBodies.Count; i++) {
                        children.Add(hybridBodies.Item(i));
                    }

                    Bodies bodies = part.Bodies;
                    for (int i = 1; i <= bodies.Count; i++) {
                        children.Add(bodies.Item(i));
                    }
                }
                // HybridBody的子节点：嵌套的HybridBodies和HybridShapes
                else if (obj is HybridBody hb) {
                    // 获取嵌套的几何图形集
                    HybridBodies nestedBodies = hb.HybridBodies;
                    for (int i = 1; i <= nestedBodies.Count; i++) {
                        children.Add(nestedBodies.Item(i));
                    }

                    // 获取几何元素

                    HybridShapes shapes = hb.HybridShapes;
                    for (int i = 1; i <= shapes.Count; i++) {
                        children.Add(shapes.Item(i));
                    }
                }
                // Body的子节点：Shapes
                else if (obj is Body body) {
                    Shapes shapes = body.Shapes;
                    for (int i = 1; i <= shapes.Count; i++) {
                        children.Add(shapes.Item(i));
                    }
                }
                // 其他类型尝试通用方法
                else {
                    // 尝试获取Children属性
                    var childrenProp = obj.GetType().GetProperty("Children");
                    if (childrenProp != null) {
                        var childrenObj = childrenProp.GetValue(obj);
                        if (childrenObj != null) {
                            var countProp = childrenObj.GetType().GetProperty("Count");
                            if (countProp != null) {
                                int count = (int)countProp.GetValue(childrenObj);
                                var itemMethod = childrenObj.GetType().GetMethod("Item");
                                if (itemMethod != null) {
                                    for (int i = 1; i <= count; i++) {
                                        var child = itemMethod.Invoke(childrenObj, new object[] { i });
                                        if (child != null) {
                                            children.Add(child);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // 添加参数节点（使用 Parameters.SubList 方法）
                // 不包括 Part 节点，因为用户不需要根节点的参数
                if (!(obj is Parameter) && !(obj is Part) && rootPart != null) {
                    try {
                        Parameters parameters = rootPart.Parameters;

                        // 尝试将对象转换为 AnyObject
                        if (obj is AnyObject anyObj) {
                            try {
                                // 使用 SubList(AnyObject, false) 获取当前对象直接拥有的参数
                                Parameters subParams = parameters.SubList(anyObj, false);

                                if (subParams != null && subParams.Count > 0) {
                                    for (int i = 1; i <= subParams.Count; i++) {
                                        Parameter param = subParams.Item(i);
                                        children.Add(param);
                                    }
                                }
                            }
                            catch {
                                // 忽略获取参数失败
                            }
                        }
                    }
                    catch {
                        // 忽略获取参数失败
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"GetChildren异常：{ex.Message}");
            }

            return children;
        }
    }
}