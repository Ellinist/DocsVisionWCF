using System;
using System.Runtime.Serialization;

namespace DocsVisionWebSide.DocsVisionClasses
{
    /// <summary>
    /// Класс отделов организации
    /// </summary>
    [DataContract]
    [Serializable]
    public class Departments
    {
        /// <summary>
        /// Идентификатор отдела
        /// </summary>
        [DataMember]
        public int IDDepartment { get; set; }

        /// <summary>
        /// Название отдела
        /// </summary>
        [DataMember]
        public string DepartmentName { get; set; }

        /// <summary>
        /// Комментарий по отделу
        /// </summary>
        [DataMember]
        public string DepartmentComment { get; set; }

        /// <summary>
        /// Флаг головного отдела
        /// Если true - отдел головной
        /// Если false - отдел обычный
        /// Допускается только один головной отдел
        /// Логика смены флага отдела выполняется на сервере (в хранимой процедуре)
        /// </summary>
        [DataMember]
        public bool MainDepartmentFlag { get; set; }
    }
}