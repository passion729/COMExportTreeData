using System.Collections.Generic;
using Newtonsoft.Json;

namespace COMExportTreeData.Models {
    /// <summary>
    /// 统一的节点数据模型 - properties作为children的一部分
    /// </summary>
    public class NodeSchema {
        [JsonProperty("name")] public string Name { get; set; }

        [JsonProperty("type", NullValueHandling = NullValueHandling.Ignore)]
        public string Type { get; set; }

        [JsonProperty("value", NullValueHandling = NullValueHandling.Ignore)]
        public string Value { get; set; }

        [JsonProperty("children", NullValueHandling = NullValueHandling.Ignore)]
        public List<NodeSchema> Children { get; set; }

        /// <summary>
        /// 判断是否应该序列化children字段
        /// </summary>
        public bool ShouldSerializeChildren() {
            return Children != null && Children.Count > 0;
        }

        public NodeSchema() {
            Children = new List<NodeSchema>();
        }

        /// <summary>
        /// 添加属性节点的辅助方法
        /// </summary>
        public void AddProperty(string name, string value) {
            Children.Add(new NodeSchema {
                Name = name,
                Value = value
            });
        }
    }
}