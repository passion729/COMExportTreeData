using System;
using System.Collections.Generic;
using KnowledgewareTypeLib;
using MECMOD;

namespace COMExportTreeData.Helpers {
    /// <summary>
    /// 参数辅助类 - 处理CATIA参数相关操作
    /// </summary>
    public static class ParameterHelper {
        /// <summary>
        /// 获取参数值
        /// </summary>
        public static string GetParameterValue(Parameter param) {
            if (param == null) return "";

            try {
                return param.ValueAsString();
            }
            catch {
                try {
                    // 尝试使用反射获取Value属性
                    var valueProperty = param.GetType().GetProperty("Value");
                    if (valueProperty != null) {
                        object paramValue = valueProperty.GetValue(param);
                        return paramValue?.ToString() ?? "";
                    }

                    return param.ToString();
                }
                catch {
                    return "";
                }
            }
        }

        /// <summary>
        /// 获取节点直接拥有的参数（不递归）
        /// </summary>
        public static List<Parameter> GetDirectParameters(object obj) {
            List<Parameter> parameters = new List<Parameter>();

            if (obj == null) return parameters;

            try {
                // 尝试获取Parameters集合
                if (obj is Part part) {
                    Parameters partParams = part.Parameters;
                    for (int i = 1; i <= partParams.Count; i++) {
                        parameters.Add(partParams.Item(i));
                    }
                }
                else if (obj is Body) {
                    // Body可能也有Parameters
                    var paramsProperty = obj.GetType().GetProperty("Parameters");
                    if (paramsProperty != null) {
                        var paramsObj = paramsProperty.GetValue(obj);
                        if (paramsObj != null) {
                            var countProperty = paramsObj.GetType().GetProperty("Count");
                            if (countProperty != null) {
                                int count = (int)countProperty.GetValue(paramsObj);
                                var itemMethod = paramsObj.GetType().GetMethod("Item");
                                if (itemMethod != null) {
                                    for (int i = 1; i <= count; i++) {
                                        var param = itemMethod.Invoke(paramsObj, new object[] { i });
                                        if (param is Parameter p) {
                                            parameters.Add(p);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else {
                    // 通用反射方法尝试获取Parameters属性
                    var paramsProperty = obj.GetType().GetProperty("Parameters");
                    if (paramsProperty != null) {
                        var paramsObj = paramsProperty.GetValue(obj);
                        if (paramsObj != null) {
                            var countProperty = paramsObj.GetType().GetProperty("Count");
                            if (countProperty != null) {
                                int count = (int)countProperty.GetValue(paramsObj);
                                var itemMethod = paramsObj.GetType().GetMethod("Item");
                                if (itemMethod != null) {
                                    for (int i = 1; i <= count; i++) {
                                        var param = itemMethod.Invoke(paramsObj, new object[] { i });
                                        if (param is Parameter p) {
                                            parameters.Add(p);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"GetDirectParameters异常：{ex.Message}");
            }

            return parameters;
        }

        /// <summary>
        /// 递归获取所有参数对象
        /// </summary>
        public static List<Parameter> GetAllParameters(object obj) {
            List<Parameter> parameters = new List<Parameter>();

            if (obj == null) return parameters;

            try {
                // 若当前对象是参数，直接添加
                if (obj is Parameter param) {
                    parameters.Add(param);
                }

                // 获取 Part 对象
                Part part = obj as Part;
                if (part == null) {
                    return parameters;
                }

                // 递归遍历子节点，收集所有参数
                List<object> children = CatiaObjectHelper.GetChildren(part, obj);
                if (children != null) {
                    foreach (var child in children) {
                        if (child is Parameter p) {
                            parameters.Add(p);
                        }
                        else {
                            parameters.AddRange(GetAllParameters(child));
                        }
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine($"GetAllParameters异常：{ex.Message}");
            }

            return parameters;
        }
    }
}