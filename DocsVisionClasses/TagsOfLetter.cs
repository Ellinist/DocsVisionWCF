using System;
using System.Runtime.Serialization;

namespace DocsVisionWebSide.DocsVisionClasses
{
    /// <summary>
    /// Класс привязки тэгов к письмам
    /// </summary>
    [DataContract]
    [Serializable]
    public class TagsOfLetter
    {
        /// <summary>
        /// Идентификатор письма, к которому привязан тэг
        /// </summary>
        [DataMember]
        public int IDLetterLink { get; set; }

        /// <summary>
        /// Идентификатор тэга, который привязан к письму
        /// </summary>
        [DataMember]
        public int IDTagLink { get; set; }

        /// <summary>
        /// Имя тэга
        /// Для логики работы ненужное свойство (!)
        /// Но, необходимо для избегания дополнительного запроса в LookUp-поле спсиска связанных тэгов
        /// </summary>
        [DataMember]
        public string TagName { get; set; }
    }
}