using System;
using System.Runtime.Serialization;

namespace DocsVisionWebSide.DocsVisionClasses
{
    /// <summary>
    /// Класс писем организации
    /// </summary>
    [DataContract]
    [Serializable]
    public class Letters
    {
        /// <summary>
        /// Идентификатор письма
        /// </summary>
        [DataMember]
        public int IDLetter { get; set; }

        /// <summary>
        /// Идентификатор отдела, к которому относится письмо
        /// </summary>
        [DataMember]
        public int IDDepartmentLetter { get; set; }

        /// <summary>
        /// Дата и время регистрации (добавления или изменения) на сервере (автоматическое задание)
        /// </summary>
        [DataMember]
        public DateTime LetterRegisterDateTime { get; set; }

        /// <summary>
        /// Название письма
        /// </summary>
        [DataMember]
        public string LetterName { get; set; }

        /// <summary>
        /// Дата получения или отправки письма
        /// </summary>
        [DataMember]
        public DateTime LetterDateTime { get; set; }

        /// <summary>
        /// Тема письма
        /// </summary>
        [DataMember]
        public string LetterTopic { get; set; }

        /// <summary>
        /// Отправитель письма
        /// </summary>
        [DataMember]
        public string LetterFrom { get; set; }

        /// <summary>
        /// Получатель (адресат) письма
        /// </summary>
        [DataMember]
        public string LetterTo { get; set; }

        /// <summary>
        /// Содержание письма (тело письма)
        /// </summary>
        [DataMember]
        public string LetterContent { get; set; }

        /// <summary>
        /// Комментарий к письму
        /// </summary>
        [DataMember]
        public string LetterComment { get; set; }

        /// <summary>
        /// Входящее (true) или исходящее (false) письмо [Флаг]
        /// </summary>
        [DataMember]
        public bool IsLetterIncoming { get; set; }
    }
}