using System;
using System.Runtime.Serialization;

namespace DocsVisionWebSide.DocsVisionClasses
{
    /// <summary>
    /// Класс тэгов
    /// </summary>
    [DataContract]
    [Serializable]
    public class Tags
    {
        /// <summary>
        /// Идентификатор тэга
        /// </summary>
        [DataMember]
        public int IDTag { get; set; }

        /// <summary>
        /// Имя тэга
        /// </summary>
        [DataMember]
        public string TagName { get; set; }
    }
}