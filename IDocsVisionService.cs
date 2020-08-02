using DocsVisionWebSide.DocsVisionClasses;
using System;
using System.Collections.Generic;
using System.ServiceModel;

namespace DocsVisionWebSide
{
    /// <summary>
    /// Определение интерфейса для работы с отделами организации
    /// </summary>
    [ServiceContract]
    public interface IDocsVisionService
    {
        /// <summary>
        /// В качестве возвращаемых значений для отделов организации используем кортежи
        /// </summary>

        /// <summary>
        /// Метод получения списка отделов организации
        /// </summary>
        /// <returns> Кортеж {Список отделов - в случае неудачи: null, Код ошибки} </returns>
        [OperationContract]
        Tuple<List<Departments>, int> GetAllDepartments();

        /// <summary>
        /// Метод добавления нового отдела
        /// </summary>
        /// <param name="departmentName">     Название отдела (string) </param>
        /// <param name="departmentComment">  Комментарий к отделу (string) </param>
        /// <param name="mainDepartmentFlag"> Флаг головного отдела (bool) {true - головной отдел, false - обычный отдел} </param>
        /// <returns> Код ошибки или индекс
        /// В случае успеха: индекс указывает на позицию добавленной записи в списке
        /// В случае неудачи: код ошибки = -1 (выборка неудачна)
        /// В случае неудачи: код ошибки = -2 (нарушение уникальности значений поля названия отдела)
        /// В случае неудачи: код ошибки = [Другой] (неизвестная ошибка - для дальнейшей разработки)
        /// </returns>
        [OperationContract]
        int AddNewDepartment(string departmentName, string departmentComment, bool mainDepartmentFlag);

        /// <summary>
        /// Метод обновления записи выбранного отдела
        /// </summary>
        /// <param name="idDepartment">       Идентификатор обновляемого отдела (int) </param>
        /// <param name="departmentName">     Название отдела (string) </param>
        /// <param name="departmentComment">  Комментарий к отделу (string) </param>
        /// <param name="mainDepartmentFlag"> Флаг головного отдела (bool) {true - головной отдел, false - обычный отдел} </param>
        /// <returns> Код ошибки или индекс
        /// В случае успеха: {код = индексу добавленной записи с связном списке}
        /// В случае неудачи: {код = код ошибки (значение меньше 0)}
        /// </returns>
        [OperationContract]
        int UpdateDepartment(int idDepartment, string departmentName, string departmentComment, bool mainDepartmentFlag);

        /// <summary>
        /// Метод удаления выбранного отдела 
        /// </summary>
        /// <param name="idDepartment"> Идентификатор удаляемого отдела (int) </param>
        /// <returns> Код ошибки
        /// В случае успеха: {код = 1}
        /// В случае неудачи: {код = код ошибки (значение меньше 0)}
        /// </returns>
        [OperationContract]
        int DeleteDepartment(int idDepartment);
               
        /// <summary>
        /// Метод получения списка писем по условию поиска
        /// Условие поиска = передаваемая строка SELECT
        /// Метод небезопасный: стоит проработать проверку в хранимой процедуре на валидность строк запросов
        /// </summary>
        /// <param name="whereParams"> Строка запроса с условием поиска </param>
        /// <returns> Кортеж
        /// Первый параметр кортежа - список писем, соответствующий условию поиска
        /// Второй параметр кортежа - код ошибки
        /// </returns>
        [OperationContract]
        Tuple<List<Letters>, int> GetLetters(string whereParams);


        /// <summary>
        /// Метод добавления нового письма
        /// </summary>
        /// <param name="idDepartmentLetter"> Идентификатор отдела, к которому относится письмо </param>
        /// <param name="letterName"> Название письма </param>
        /// <param name="letterDateTime"> Дата получения или отправки письма </param>
        /// <param name="letterTopic"> Тема письма </param>
        /// <param name="letterFrom"> Отправитель письма </param>
        /// <param name="letterTo"> Получатель письма (адресат) </param>
        /// <param name="letterContent"> Содержание письма (тело письма) </param>
        /// <param name="letterComment"> Комментарий к письму </param>
        /// <param name="isLetterIncoming"> Флаг: входящее или исходящее письмо (входящее: true, исходящее: false) </param>
        /// <param name="whereCondition"> Критерий (условие) поиска, заданное в окне работы с письмами </param>
        /// <returns> Кортеж
        /// Первый параметр = код ошибки или индекс добавленной записи в списке
        /// в случае, если добавленное письмо не будет соответствовать критерию (условию) поиска, равно 100000000 (сто миллионов)
        /// требуется для уведомления пользователя о том, что письмо создано, но отображено не будет
        /// логика не доработана - необходимо продумать иной вариант
        /// Второй параметр - идентификатор добавленной записи
        /// </returns>
        [OperationContract]
        Tuple<int, int> AddLetter(int idDepartmentLetter, string letterName, DateTime letterDateTime, string letterTopic, string letterFrom, string letterTo,
                                  string letterContent, string letterComment, bool isLetterIncoming, string whereCondition);

        /// <summary>
        /// Метод удаления письма
        /// </summary>
        /// <param name="idLetter"> Идентификатор удяляемого письма </param>
        /// <returns> Код ошибки </returns>
        [OperationContract]
        int DeleteLetter(int idLetter);

        /// <summary>
        /// Метод редактирования выбранного письма
        /// </summary>
        /// <param name="idLetter"> Идентификатор редактируемого письма </param>
        /// <param name="idDepartmentLetter"> Идентификатор отдела, к которому привязано письмо </param>
        /// <param name="letterName"> Название письма </param>
        /// <param name="letterDateTime"> Дата получения ил отправки письма </param>
        /// <param name="letterTopic"> Тема письма </param>
        /// <param name="letterFrom"> Отправитель письма </param>
        /// <param name="letterTo"> Получатель письма (адресат) </param>
        /// <param name="letterContent"> Содержание письма (тело письма) </param>
        /// <param name="letterComment"> Комментарий к письму </param>
        /// <param name="isLetterIncoming"> Флаг: входящее или исходящее письмо (входящее: true, исходящее: false) </param>
        /// <param name="whereCondition"> Критерий (условие) поиска, заданное в окне работы с письмами </param>
        /// <returns> Код ошибки или индекс
        /// в случае, если добавленное письмо не будет соответствовать критерию (условию) поиска, равно 100000000 (сто миллионов)
        /// требуется для уведомления пользователя о том, что письмо создано, но отображено не будет
        /// логика не доработана - необходимо продумать иной вариант 
        /// </returns>
        [OperationContract]
        int EditLetter(int idLetter, int idDepartmentLetter, string letterName, DateTime letterDateTime, string letterTopic, string letterFrom, string letterTo,
                       string letterContent, string letterComment, bool isLetterIncoming, string whereCondition);

        /// <summary>
        /// Метод получения списка тэгов конкретного письма
        /// </summary>
        /// <param name="idLetter"> Идентификатор выбранного письма </param>
        /// <returns> Кортеж
        /// Первый параметр = список тэгов выбранного письма
        /// Второй параметр = код ошибки
        /// </returns>
        [OperationContract]
        Tuple<List<TagsOfLetter>, int> GetTagsOfLetter(int idLetter);

        /// <summary>
        /// Метод обновления списка тэгов выбранного письма (после корректировки)
        /// </summary>
        /// <param name="tagsOfLetterList"> Список обновившегося набора тэго письма </param>
        /// <param name="idLetter"> Идентификатор письма, для которого проведены изменения </param>
        /// <returns> Код ошибки </returns>
        [OperationContract]
        int UpdateTagsOfLetter(List<TagsOfLetter> tagsOfLetterList, int idLetter);

        /// <summary>
        /// Метод получения списка существующих в БД тэгов
        /// </summary>
        /// <returns> Кортеж
        /// Первый параметр = список тэгов
        /// Второй параметр = код ошибки
        /// </returns>
        [OperationContract]
        Tuple<List<Tags>, int> GetTags();

        /// <summary>
        /// Метод удаления тэга из списка тэгов
        /// Внимание! Удаление тэга, который связан отношение многие-ко-многим с таблицей писем,
        /// приведет к удалению этого тэга из промежуточной таблицы.
        /// Данный тэг станет недоступным для всех писем (или удалится у них)
        /// </summary>
        /// <param name="idTag"> Идентификатор удаляемого тэга </param>
        /// <returns> Код ошибки при удалении </returns>
        [OperationContract]
        int DeleteTag(int idTag);

        /// <summary>
        /// Метод добавления нового тэга в список доступных тэгов
        /// </summary>
        /// <param name="tag"> Создаваемый тэг </param>
        /// <returns> Код ошибки при добавлении </returns>
        [OperationContract]
        int AddTag(string tag);

        /// <summary>
        /// Метод переименования тэга
        /// </summary>
        /// <param name="idTag"> Идентификатор переименовываемого тэга </param>
        /// <param name="tag"> Новое название тэга </param>
        /// <returns> Код ошибки при переименовании </returns>
        [OperationContract]
        int RenameTag(int idTag, string tag);
    }
}
