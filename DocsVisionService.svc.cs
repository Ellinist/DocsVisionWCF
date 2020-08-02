using DocsVisionWebSide.DocsVisionClasses;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Runtime.Serialization;

namespace DocsVisionWebSide
{
    /// <summary>
    /// Класс логики работы с отделами (выборка, запись, обновление, удаление - CRUD)
    /// </summary>
    [Serializable]
    [DataContract]
    public class DocsVisionService : IDocsVisionService, IDisposable
    {
        #region Соединение с БД
        private readonly string connectionDocsVisionString; // Строка соединения
        SqlConnection connectionDocsVision;                 // Соединение
        #endregion

        #region Список хранимых процедур на сервере
        // Пока хранимые процедуры - потом сделать список контрактов (после добавления Web-сервисов
        private readonly string selectDepartmentsString = "DocsVisionSelectDepartments"; // Строка, идентифицирующая сохраненную процедуру выборки записей на сервере
        private readonly string insertDepartmentString  = "DocsVisionInsertDepartment";  // Строка, идентифицирующая сохраненную процедуру добавления записи на сервере
        private readonly string deleteDepartmentString  = "DocsVisionDeleteDepartment";  // Строка, идентифицирующая сохраненную процедуру удаления записи на сервере
        private readonly string updateDepartmentString  = "DocsVisionUpdateDepartment";  // Строка, идентифицирующая сохраненную процедуру обновления записи на сервере
        private readonly string selectLettersString     = "DocsVisionSelectLetters";     // Строка, идентифицирующая сохраненную процедуру выборки записей писем на сервере
        private readonly string insertLetterString      = "DocsVisionInsertLetter";      // Строка, идентифицирующая сохраненную процедуру добавления нового письма на сервере
        private readonly string deleteLetterString      = "DocsVisionDeleteLetter";      // Строка, идентифицирующая сохраненную процедуру удаления письма на сервере
        private readonly string updateLetterString      = "DocsVisionUpdateLetter";      // Строка, идентифицирующая сохраненную процедуру обновления письма на сервере
        private readonly string selectTagsString        = "DocsVisionSelectTags";        // Строка, идентифицирующая сохраненную процедуру выборки тэгов на сервере
        private readonly string deleteTagString         = "DocsVisionDeleteTag";         // Строка, идентифицирующая сохраненную процедуру удаления тэга на сервере
        private readonly string addTagString            = "DocsVisionAddTag";            // Строка, идентифицирующая сохраненную процедуру добавления тэга на сервере
        private readonly string renameTagString         = "DocsVisionRenameTag";         // Строка, идентифицирующая сохраненную процедуру переименования тэга на сервере
        #endregion

        /// <summary>
        /// Конструктор класса DocsVisionService
        /// </summary>
        public DocsVisionService()
        {
            // Определение соединения с СУБД
            connectionDocsVisionString = ConfigurationManager.ConnectionStrings["DocsVisionConnectionString"].ConnectionString;
        }

        /// <summary>
        /// Метод получения списка отделов организации
        /// </summary>
        /// <returns> Кортеж {Связный список отделов - в случае неудачи: null, Код ошибки} </returns>
        public Tuple<List<Departments>, int> GetAllDepartments()
        {
            int selectErrNo = 1; // Этот код говорит о том, что ошибки нет (если все пройдет успешно, значение не изменится)
            List<Departments> selectDepartmentsList = new List<Departments>(); // Создаем новый связный список отделов
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString)) // Создаем экземпляр соединения
            {
                #region Определение команды для выборки списка отделов из БД
                SqlCommand selectCommand = new SqlCommand()      // Создаем SQL команду
                {
                    Connection  = connectionDocsVision,          // Определяем соединение
                    CommandText = selectDepartmentsString,       // Определяем строку хранимой процедуры SELECT
                    CommandType = CommandType.StoredProcedure    // Задаем параметр применения хранимой процедуры
                    // Так как процедура не принимает никаких параметров - ограничиваемся только этим
                };
                #endregion
                #region Блок try-catch-finally
                try
                {
                    connectionDocsVision.Open();  // Открываем соединение
                    SqlDataReader readerDocsVisionDepartments = selectCommand.ExecuteReader(); // Читаем из БД в Reader
                    // В цикле проходим по полученным записям для заполнения связного списка отделов
                    foreach (DbDataRecord dr in readerDocsVisionDepartments)
                    {
                        selectDepartmentsList.Add(new Departments() // Добавляем новую запись в связный список
                        {
                            // Формируем значения свойств класса отделов (Departments)
                            IDDepartment = Int32.Parse(dr["id_Department"].ToString()), // Идентификатор отдела (задается сервером - автоинкремент)
                            DepartmentName = dr["departmentName"].ToString(),           // Название отдела
                            DepartmentComment = dr["departmentComment"].ToString(),     // Комментарий по отделу
                            MainDepartmentFlag = (bool)dr["mainDepartmentFlag"]         // Флаг головного отдела (true - головной, false - обычный)
                        });
                    }
                    connectionDocsVision.Close(); // Закрываем соединение
                }
                catch(Exception selectException)
                {
                    selectErrNo = selectException.HResult; // Этот код говорит о том, что произошла ошибка при выборке записей из БД
                    selectDepartmentsList = null; // Опустошаем связный список
                }
                finally
                {
                    connectionDocsVision.Dispose(); // Уничтожаем соединение
                }
                #endregion
            }
            // В случае успеха: код ошибки = 1, заполненный связный список
            // В случае неудачи: код ошибки = -1, связный список = null
            return new Tuple<List<Departments>, int>(selectDepartmentsList, selectErrNo); // Возвращаем кортеж-фигешь
        }

        /// <summary>
        /// Метод добавления нового отдела
        /// </summary>
        /// <param name="departmentName">     Название отдела (string) </param>
        /// <param name="departmentComment">  Комментарий к отделу (string) </param>
        /// <param name="mainDepartmentFlag"> Флаг головного отдела (bool) {true - головной отдел, false - обычный отдел} </param>
        /// <returns> Код ошибки или индекс
        /// В случае успеха: {код = индексу добавленной записи с связном списке, заполненный связный список}
        /// В случае неудачи: {код = код ошибки (значение меньше 0), связный список = null}
        /// </returns>
        public int AddNewDepartment(string departmentName, string departmentComment, bool mainDepartmentFlag)
        {
            int index = -1; // Этот код говорит о том, что возникла ошибка (если все пройдет успешно, значение поменяется на индекс записи)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString)) // Создаем экземпляр соединения
            {
                #region Определение команды для добавления нового отдела
                SqlCommand insertCommand = new SqlCommand() // Создаем SQL команду
                {
                    Connection = connectionDocsVision,           // Определяем соединение
                    CommandText = insertDepartmentString,        // Определяем строку хранимой процедуры INSERT
                    CommandType = CommandType.StoredProcedure    // Задаем параметр применения хранимой процедуры
                };
                #endregion
                #region Задание параметров в процедуру DocsVisionInsertDepartment
                SqlParameter insertParam_1 = new SqlParameter // Первый параметр - название отдела организации (параметр входной)
                {
                    ParameterName = "@DEPARTMENTNAME",
                    Value = departmentName,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_1);
                SqlParameter insertParam_2 = new SqlParameter // Второй параметр - комментарий по отделу (параметр входной)
                {
                    ParameterName = "@DEPARTMENTCOMMENT",
                    Value = departmentComment,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_2);
                SqlParameter insertParam_3 = new SqlParameter // Третий параметр - флаг головного отдела (параметр входной)
                {
                    ParameterName = "@MAINDEPARTMENTFLAG",
                    Value = mainDepartmentFlag,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_3);
                SqlParameter returnParam = new SqlParameter // Четвертый параметр - идентификатор последней занесенной записи (параметр выходной)
                {
                    ParameterName = "@LASTRECORDID",
                    Value = 0,
                    Direction = ParameterDirection.Output
                };
                insertCommand.Parameters.Add(returnParam);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    #region Блок выполнения процедуры добавления записи
                    connectionDocsVision.Open();     // Открываем соединение
                    insertCommand.ExecuteNonQuery(); // Выполняем команду добавления записи
                    connectionDocsVision.Close();    // Закрываем соединение
                    #endregion
                    if (GetAllDepartments().Item2 != -1)
                    {
                        // Выборка прошла успешно
                        int i = 0; // Создаем счетчик - для определения индекса записи
                        foreach (Departments dep in GetAllDepartments().Item1) // В цикле ищем индекс добавленной записи
                        {
                            if (dep.IDDepartment == (int)returnParam.Value)
                            {
                                index = i; // В случае совпадения ID нового отдела и позиции в списке - индекс получает номер позиции
                                break;     // Прерываем цикл
                            }
                            i++; // Увеличиваем счетчик
                        }
                    }
                    else
                    {
                        // В случае неудачной выборки в таблице отделов задаем флаг ошибки
                        index = -1; // Выборка не осуществилась
                    }
                }
                catch (Exception insertException)
                {
                    if (insertException.HResult == -2146232060) // Код ошибки - нарушение правила индекса unique
                    {
                        index = -2; // Данное возвращаемое значение говорит о нарушении индекса уникальности названия отдела
                    }
                    else
                    {
                        index = insertException.HResult; // Неопознанный код ошибки - для дальнейших разборов
                    }
                }
                finally
                {
                    connectionDocsVision.Dispose(); // Уничтожаем соединение
                }
                #endregion
            }
            // В случае успеха: индекс указывается на позицию добавленной записи в списке
            // В случае неудачи: код ошибки = -1 (выборка неудачна)
            // В случае неудачи: код ошибки = -2 (нарушение уникальности значений поля названия отдела)
            // В случае неудачи: код ошибки = [Другой] (неизвестная ошибка - для дальнейшей разработки)
            return index;
        }

        /// <summary>
        /// Метод удаления выбранного отдела 
        /// </summary>
        /// <param name="idDepartment"> Идентификатор удаляемого отдела (int) </param>
        /// <returns> Код ошибки
        /// В случае успеха: {код = 1}
        /// В случае неудачи: {код = код ошибки (значение меньше 0)}
        /// </returns>
        public int DeleteDepartment(int idDepartment)
        {
            int selectErrNo = 1; // Этот код говорит о том, что ошибки нет (если все пройдет успешно, значение не изменится)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString)) // Создаем экземпляр соединения
            {
                #region Определение команды для удаления выбранного отдела
                SqlCommand deleteCommand = new SqlCommand() // Создаем SQL команду
                {
                    Connection = connectionDocsVision,           // Определяем соединение
                    CommandText = deleteDepartmentString,        // Определяем строку хранимой процедуры DELETE
                    CommandType = CommandType.StoredProcedure    // Задаем параметр применения хранимой процедуры
                };
                #endregion
                #region Задание параметра для процедуры DocsVisionDeleteDepartment
                SqlParameter deleteParam = new SqlParameter()
                {
                    ParameterName = "@IDDELETEDEPARTMENT",
                    Value = idDepartment,
                    Direction = ParameterDirection.Input
                };
                deleteCommand.Parameters.Add(deleteParam);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    #region Блок выполнения команды удаления записи
                    connectionDocsVision.Open();     // Открываем соединение
                    deleteCommand.ExecuteNonQuery(); // Выполняем команду удаления выбранной записи
                    connectionDocsVision.Close();    // Закрываем соединение
                    #endregion
                }
                catch(Exception deleteException)
                {
                    if(deleteException.HResult == -2146232060)
                    {
                        // Попытка удалить головной отдел
                        selectErrNo = -2;
                    }
                    else
                    {
                        selectErrNo = deleteException.HResult; // Произошла ошибка в процедуре удаления записи
                    }
                }
                finally
                {
                    connectionDocsVision.Dispose(); // Уничтожаем соединение
                }
                #endregion
            }
            return selectErrNo; // Возвращаем флаг результата выполнения операции удаления записи
        }

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
        public int UpdateDepartment(int idDepartment, string departmentName, string departmentComment, bool mainDepartmentFlag)
        {
            int index = -1; // Этот код говорит о том, что возникла ошибка (если все пройдет успешно, значение поменяется на индекс записи)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString)) // Создаем экземпляр соединения
            {
                #region Определение команды для обновления отдела
                SqlCommand updateCommand = new SqlCommand()             // Создаем SQL команду
                {
                    Connection = connectionDocsVision,         // Определяем соединение
                    CommandText = updateDepartmentString,      // Определяем строку хранимой процедуры UPDATE
                    CommandType = CommandType.StoredProcedure // Задаем параметр применения хранимой процедуры
                };
                    
                #endregion
                #region Задание параметров в процедуру DocsVisionUpdateDepartment
                SqlParameter updateParam_1 = new SqlParameter // Первый параметр - идентификатор обновляемого отдела (параметр входной)
                {
                    ParameterName = "@IDUPDATEDEPARTMENT",
                    Value = idDepartment,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(updateParam_1);
                SqlParameter updateParam_2 = new SqlParameter // Второй параметр - название обновляемого отдела (параметр входной)
                {
                    ParameterName = "@DEPARTMENTNAME",
                    Value = departmentName,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(updateParam_2);
                SqlParameter updateParam_3 = new SqlParameter // Третий параметр - комментарий по обновляемому отделу (параметр входной)
                {
                    ParameterName = "@DEPARTMENTCOMMENT",
                    Value = departmentComment,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(updateParam_3);
                SqlParameter updateParam_4 = new SqlParameter // Четвертый параметр - Флаг, головной ли обновляемый отдел (параметр входной)
                {
                    ParameterName = "@MAINDEPARTMENTFLAG",
                    Value = mainDepartmentFlag,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(updateParam_4);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    #region Блок выполнения процедуры обновления записи
                    connectionDocsVision.Open();          // Открываем соединение
                    updateCommand.ExecuteNonQuery(); // Выполняем команду обновления записи
                    connectionDocsVision.Close();         // Закрываем соединение
                    #endregion
                    if (GetAllDepartments().Item2 != -1)
                    {
                        // Выборка прошла успешно
                        int i = 0; // Создаем счетчик - для определения индекса записи
                        foreach (Departments dep in GetAllDepartments().Item1) // В цикле ищем индекс обновленной записи
                        {
                            if (dep.IDDepartment == idDepartment)
                            {
                                index = i; // В случае совпадения ID обновленного отдела и позиции в списке - индекс получает номер позиции
                                break;     // Прерываем цикл
                            }
                            i++; // Увеличиваем счетчик
                        }
                    }
                    else
                    {
                        // В случае неудачной выборки в таблице отделов задаем флаг ошибки
                        index = -1; // Выборка не осуществилась
                    }
                }
                catch (Exception updateException)
                {
                    if (updateException.HResult == -2146232060) // Код ошибки - нарушение правила индекса unique
                    {
                        index = -2; // Данное возвращаемое значение говорит о нарушении индекса уникальности названия отдела
                    }
                    else
                    {
                        index = updateException.HResult; // Неопознанный код ошибки - для дальнейших разборов
                    }
                }
                finally
                {
                    connectionDocsVision.Dispose();
                }
                #endregion
            }
            // В случае успеха: индекс указывается на позицию обновленной записи в списке
            // В случае неудачи: код ошибки = -1 (выборка неудачна)
            // В случае неудачи: код ошибки = -2 (нарушение уникальности значений поля названия отдела)
            // В случае неудачи: код ошибки = [Другой] (неизвестная ошибка - для дальнейшей разработки)
            return index;
        }

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
        public Tuple<List<Letters>, int> GetLetters(string whereParams)
        {
            List<Letters> lettersList = new List<Letters>();
            int selectErrNo = 1; // Этот код говорит о том, что ошибки нет (если все пройдет успешно, значение не поменяется)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString))
            {
                #region Определение команды для выборки списка писем из БД
                SqlCommand selectCommand = new SqlCommand()   // Создаем SQL команду
                {
                    Connection  = connectionDocsVision,       // Определяем соединение
                    CommandText = selectLettersString,  // Определяем строку хранимой процедуры SELECT
                    CommandType = CommandType.StoredProcedure // Задаем параметр применения хранимой процедуры
                };
                #endregion
                #region Задание параметра поиска списка писем
                SqlParameter whereParam = new SqlParameter   // Определяем параметр строки поиска
                {
                    ParameterName = "@WHERECONDITION",
                    Value = whereParams,
                    Direction = ParameterDirection.Input
                };
                selectCommand.Parameters.Add(whereParam);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    #region Блок получения списка писем с сервера
                    connectionDocsVision.Open(); // Устанавливаем соединение
                    SqlDataReader selectReader = selectCommand.ExecuteReader(); // Получае Reader
                    foreach (DbDataRecord dr in selectReader) // Запускаем цикл заполнения списка писем
                    {
                        lettersList.Add(new Letters()
                        {
                            IDLetter = Int32.Parse(dr["id_Letter"].ToString()),
                            IDDepartmentLetter = Int32.Parse(dr["id_DepartmentLetter"].ToString()),
                            LetterRegisterDateTime = (DateTime)dr["letterRegisterDateTime"],
                            LetterName = dr["letterName"].ToString(),
                            LetterDateTime = (DateTime)dr["letterDateTime"],
                            LetterTopic = dr["letterTopic"].ToString(),
                            LetterFrom = dr["letterFrom"].ToString(),
                            LetterTo = dr["letterTo"].ToString(),
                            LetterContent = dr["letterContent"].ToString(),
                            LetterComment = dr["letterComment"].ToString(),
                            IsLetterIncoming = (bool)dr["isLetterIncoming"]
                        });
                    }
                    connectionDocsVision.Close(); // Закрываем соединение
                    #endregion
                }
                catch (Exception selectException)
                {
                    selectErrNo = selectException.HResult; // При возникновении ошибки - возвращаем код ошибки
                    lettersList = null;
                }
                finally
                {
                    connectionDocsVision.Dispose();
                }
                #endregion
            }
            return new Tuple<List<Letters>, int>(lettersList, selectErrNo);
        }

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
        public Tuple<int, int> AddLetter(int idDepartmentLetter, string letterName, DateTime letterDateTime, string letterTopic, string letterFrom, string letterTo,
                                         string letterContent, string letterComment, bool isLetterIncoming, string whereCondition)
        {
            List<Letters> lettersList; // Объявляем новый список
            int index = -1; // Этот код говорит о том, что возникла ошибка (если все пройдет успешно, значение поменяется на индекс записи)
            SqlParameter returnParam;

            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString))
            {
                #region Определение команды для добавления нового письма
                SqlCommand insertCommand = new SqlCommand()   // Создаем SQL команду
                {
                    Connection  = connectionDocsVision,       // Определяем соединение
                    CommandText = insertLetterString,         // Определяем строку хранимой процедуры INSERT
                    CommandType = CommandType.StoredProcedure // Задаем параметр применения хранимой процедуры
                };
                #endregion
                #region Задание параметров в процедуру DocsVisionInsertLetter
                SqlParameter insertParam_1 = new SqlParameter // Первый параметр - идентификатор связанного с письмом отдела (параметр входной)
                {
                    ParameterName = "@IDDEPARTMENTLETTER",
                    Value = idDepartmentLetter,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_1);
                SqlParameter insertParam_2 = new SqlParameter // Второй параметр - название письма (параметр входной)
                {
                    ParameterName = "@LETTERNAME",
                    Value = letterName,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_2);
                SqlParameter insertParam_3 = new SqlParameter // Третий параметр - дата получения или отправки письма (параметр входной)
                {
                    ParameterName = "@LETTERDATETIME",
                    Value = letterDateTime,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_3);
                SqlParameter insertParam_4 = new SqlParameter // Четвертый параметр - тема письма (параметр входной)
                {
                    ParameterName = "@LETTERTOPIC",
                    Value = letterTopic,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_4);
                SqlParameter insertParam_5 = new SqlParameter // Пятый параметр - отправитель (параметр входной)
                {
                    ParameterName = "@LETTERFROM",
                    Value = letterFrom,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_5);
                SqlParameter insertParam_6 = new SqlParameter // Шестой параметр - адресат (параметр входной)
                {
                    ParameterName = "@LETTERTO",
                    Value = letterTo,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_6);
                SqlParameter insertParam_7 = new SqlParameter // Седьмой параметр - содержимое письма (параметр входной)
                {
                    ParameterName = "@LETTERCONTENT",
                    Value = letterContent,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_7);
                SqlParameter insertParam_8 = new SqlParameter // Восьмой параметр - комментарий к письму (параметр входной)
                {
                    ParameterName = "@LETTERCOMMENT",
                    Value = letterComment,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_8);
                SqlParameter insertParam_9 = new SqlParameter // Девятый параметр - флаг: входящее (true) или исходящее (false) письмо (параметр входной)
                {
                    ParameterName = "@ISLETTERINCOMING",
                    Value = isLetterIncoming,
                    Direction = ParameterDirection.Input
                };
                insertCommand.Parameters.Add(insertParam_9);
                returnParam = new SqlParameter                // Десятый параметр - идентификатор последней занесенной записи (параметр выходной)
                {
                    ParameterName = "@LASTRECORDID",
                    Value = 0,
                    Direction = ParameterDirection.Output
                };
                insertCommand.Parameters.Add(returnParam);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    #region Блок выполнения процедуры добавления записи (новое письмо)
                    connectionDocsVision.Open();      // Открываем соединение
                    insertCommand.ExecuteNonQuery();  // Выполняем команду добавления записи
                    connectionDocsVision.Close();     // Закрываем соединение
                    #endregion
                    lettersList = GetLetters(whereCondition).Item1;
                    int i = 0; // Создаем счетчик - для определения индекса записи
                    foreach (Letters ltr in lettersList) // В цикле ищем индекс добавленной записи
                    {
                        if (ltr.IDLetter == (int)returnParam.Value)
                        {
                            index = i; // В случае совпадения ID нового отдела и позиции в списке - индекс получает номер позиции
                            break;     // Прерываем цикл
                        }
                        i++; // Увеличиваем счетчик
                        index = 100000000; // В выборке по указанному динамическому запросу индекс не найден - в клиенте позиционируемся на первую запись выборки
                    }
                }
                catch (Exception insertException)
                {
                    if (insertException.HResult == -2146232060) // Код ошибки - нарушение правила индекса unique
                    {
                        index = -2; // Данное возвращаемое значение говорит о нарушении индекса уникальности названия отдела
                    }
                    else
                    {
                        index = insertException.HResult; // Неопознанный код ошибки - для дальнейших разборов
                    }
                }
                finally
                {
                    connectionDocsVision.Dispose();
                }
                #endregion
            }
            return new Tuple<int, int>(index, (int)returnParam.Value);
        }

        /// <summary>
        /// Метод удаления письма
        /// </summary>
        /// <param name="idLetter"> Идентификатор удяляемого письма </param>
        /// <returns> Код ошибки </returns>
        public int DeleteLetter(int idLetter)
        {
            int deleteErrNo = 1; // Этот код говорит о том, что ошибок нет (если все пройдет успешно, код не изменится)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString))
            {
                #region Определение команды для удаления выбранного письма
                SqlCommand deleteCommand = new SqlCommand()  // Создаем SQL команду
                {
                    Connection  = connectionDocsVision,       // Определяем соединение
                    CommandText = deleteLetterString,         // Определяем строку хранимой процедуры DELETE
                    CommandType = CommandType.StoredProcedure // Задаем параметр применения хранимой процедуры
                };
                #endregion
                #region Задание параметра для процедуры DocsVisionDeleteLetter
                SqlParameter deleteParam = new SqlParameter()
                {
                    ParameterName = "@IDDELETELETTER",
                    Value = idLetter,
                    Direction = ParameterDirection.Input
                };
                deleteCommand.Parameters.Add(deleteParam);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    #region Блок выполнения команды удаления записи
                    connectionDocsVision.Open();          // Открываем соединение
                    deleteCommand.ExecuteNonQuery(); // Выполняем команду удаления выбранной записи
                    connectionDocsVision.Close();         // Закрываем соединение
                    #endregion
                }
                catch (Exception deleteException)
                {
                    deleteErrNo = deleteException.HResult; // Возвращаем код ошибки
                }
                finally
                {
                    connectionDocsVision.Dispose(); // Уничтожаем соединение
                }
                #endregion
            }
            return deleteErrNo;
        }

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
        public int EditLetter(int idLetter, int idDepartmentLetter, string letterName, DateTime letterDateTime, string letterTopic, string letterFrom, string letterTo,
                              string letterContent, string letterComment, bool isLetterIncoming, string whereCondition)
        {
            List<Letters> lettersList; // Объявляем новый список
            int index = -1; // Этот код говорит о том, что возникла ошибка (если все пройдет успешно, значение поменяется на индекс записи)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString))
            {
                #region Определение команды для добавления нового письма
                SqlCommand updateCommand = new SqlCommand()   // Создаем SQL команду
                {
                    Connection  = connectionDocsVision,       // Определяем соединение
                    CommandText = updateLetterString,         // Определяем строку хранимой процедуры INSERT
                    CommandType = CommandType.StoredProcedure // Задаем параметр применения хранимой процедуры
                };
                #endregion
                #region Задание параметров в процедуру DocsVisionUpdateLetter
                SqlParameter idParam = new SqlParameter       // Параметр ID - идентификатор обновляемого письма (параметр входной)
                {
                    ParameterName = "@IDLETTER",
                    Value = idLetter,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(idParam);
                SqlParameter insertParam_1 = new SqlParameter // Первый параметр - идентификатор связанного с письмом отдела (параметр входной)
                {
                    ParameterName = "@IDDEPARTMENTLETTER",
                    Value = idDepartmentLetter,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_1);
                SqlParameter insertParam_2 = new SqlParameter // Второй параметр - название письма (параметр входной)
                {
                    ParameterName = "@LETTERNAME",
                    Value = letterName,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_2);
                SqlParameter insertParam_3 = new SqlParameter // Третий параметр - дата получения или отправки письма (параметр входной)
                {
                    ParameterName = "@LETTERDATETIME",
                    Value = letterDateTime,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_3);
                SqlParameter insertParam_4 = new SqlParameter // Четвертый параметр - тема письма (параметр входной)
                {
                    ParameterName = "@LETTERTOPIC",
                    Value = letterTopic,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_4);
                SqlParameter insertParam_5 = new SqlParameter // Пятый параметр - отправитель (параметр входной)
                {
                    ParameterName = "@LETTERFROM",
                    Value = letterFrom,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_5);
                SqlParameter insertParam_6 = new SqlParameter // Шестой параметр - адресат (параметр входной)
                {
                    ParameterName = "@LETTERTO",
                    Value = letterTo,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_6);
                SqlParameter insertParam_7 = new SqlParameter // Седьмой параметр - содержимое письма (параметр входной)
                {
                    ParameterName = "@LETTERCONTENT",
                    Value = letterContent,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_7);
                SqlParameter insertParam_8 = new SqlParameter // Восьмой параметр - комментарий к письму (параметр входной)
                {
                    ParameterName = "@LETTERCOMMENT",
                    Value = letterComment,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_8);
                SqlParameter insertParam_9 = new SqlParameter // Девятый параметр - флаг: входящее (true) или исходящее (false) письмо (параметр входной)
                {
                    ParameterName = "@ISLETTERINCOMING",
                    Value = isLetterIncoming,
                    Direction = ParameterDirection.Input
                };
                updateCommand.Parameters.Add(insertParam_9);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    #region Блок выполнения процедуры обновления записи (редактируемое письмо)
                    connectionDocsVision.Open();     // Открываем соединение
                    updateCommand.ExecuteNonQuery(); // Выполняем команду обновления записи
                    connectionDocsVision.Close();    // Закрываем соединение
                    #endregion
                    lettersList = GetLetters(whereCondition).Item1;
                    int i = 0; // Создаем счетчик - для определения индекса записи
                    foreach (Letters ltr in lettersList) // В цикле ищем индекс добавленной записи
                    {
                        if (ltr.IDLetter == idLetter)
                        {
                            index = i; // В случае совпадения ID нового отдела и позиции в списке - индекс получает номер позиции
                            break;     // Прерываем цикл
                        }
                        i++; // Увеличиваем счетчик
                        index = 100000000; // В выборке по указанному динамическому запросу индекс не найден - в клиенте позиционируемся на первую запись выборки
                    }
                }
                catch (Exception updateException)
                {
                    if (updateException.HResult == -2146232060) // Код ошибки - нарушение правила индекса unique
                    {
                        index = -2; // Данное возвращаемое значение говорит о нарушении индекса уникальности названия отдела
                    }
                    else
                    {
                        index = updateException.HResult; // Неопознанный код ошибки - для дальнейших разборов
                    }
                }
                finally
                {
                    connectionDocsVision.Dispose();
                }
                #endregion
            }
            return index;
        }

        /// <summary>
        /// Метод получения списка тэгов конкретного письма
        /// </summary>
        /// <param name="idLetter"> Идентификатор выбранного письма </param>
        /// <returns> Кортеж
        /// Первый параметр = список тэгов выбранного письма
        /// Второй параметр = код ошибки
        /// </returns>
        public Tuple<List<TagsOfLetter>, int> GetTagsOfLetter(int idLetter)
        {
            int tagsErrNo = 1; // Этот код говорит о том, что ошибки нет (если все пройдет успешно, код не изменится)
            // Создаем строку запроса по выбранному письму
            string tagsOfLetterString = $"SELECT id_LetterLink, id_TagLink, TagName FROM tbTagsOfLetters, tbTags WHERE (id_TagLink = id_Tag AND id_LetterLink = '{idLetter}')";
            List<TagsOfLetter> tagsOfLetterList = new List<TagsOfLetter>();
            SqlDataAdapter tagsOfLetterAdapter = new SqlDataAdapter(tagsOfLetterString, connectionDocsVisionString); // Создаем адаптер
            DataSet tagsOfLetterDS = new DataSet();   // Создаем DataSet
            #region Блок try-catch-finally
            try
            {
                tagsOfLetterAdapter.Fill(tagsOfLetterDS); // Заполняем DataSet (по-сути - привязка)
                foreach (DataRow dr in tagsOfLetterDS.Tables[0].Rows) // Заполняем список в цикле по DataTable нашего DataSet
                {
                    tagsOfLetterList.Add(new TagsOfLetter()
                    {
                        IDLetterLink = Int32.Parse(dr["id_LetterLink"].ToString()),
                        IDTagLink = Int32.Parse(dr["id_TagLink"].ToString()),
                        TagName = dr["TagName"].ToString()
                    });
                }
            }
            catch(Exception tagsException)
            {
                tagsErrNo = tagsException.HResult; // Возвращаем код ошибки
            }
            finally
            {
                tagsOfLetterAdapter.Dispose();
            }
            #endregion
            return new Tuple<List<TagsOfLetter>, int>(tagsOfLetterList, tagsErrNo);
        }

        /// <summary>
        /// Метод обновления списка тэгов выбранного письма (после корректировки)
        /// </summary>
        /// <param name="tagsOfLetterList"> Список обновившегося набора тэго письма </param>
        /// <param name="idLetter"> Идентификатор письма, для которого проведены изменения </param>
        /// <returns> Код ошибки </returns>
        public int UpdateTagsOfLetter(List<TagsOfLetter> tagsOfLetterList, int idLetter)
        {
            // В этом методе происходит работа с адаптерами и построителями
            int tagsErrNo = 1; // Этот код говори о том, что ошибки нет (если все пройдет успешно, код не изменится)
            string tagsOfLetterString = $"SELECT * FROM tbTagsOfLetters WHERE id_LetterLink = '{idLetter}' AND id_Link >= 0";
            SqlDataAdapter tagsOfLetterAdapter = new SqlDataAdapter(tagsOfLetterString, connectionDocsVisionString);
            DataSet tagsOfLetterDS = new DataSet();
            #region Блок try-catch-finally
            try
            {
                tagsOfLetterAdapter.Fill(tagsOfLetterDS);
                SqlCommandBuilder builder = new SqlCommandBuilder(tagsOfLetterAdapter);
                tagsOfLetterAdapter.DeleteCommand = builder.GetDeleteCommand();
                tagsOfLetterAdapter.UpdateCommand = builder.GetUpdateCommand();
                tagsOfLetterAdapter.InsertCommand = builder.GetInsertCommand();
                // Очищаем старый набор тэгов, относящихся к выбранному письму
                // Плохой подход - необходимо придумать другой метод, отслеживающий статус RowState у объекта DataTable
                // Пока оставляю так
                foreach (DataRow dr in tagsOfLetterDS.Tables[0].Rows)
                {
                    dr.Delete();
                }
                // В цикле заносим новые значения тэгов, относящихся к выбранному письму
                foreach (TagsOfLetter tags in tagsOfLetterList)
                {
                    DataRow newRow = tagsOfLetterDS.Tables[0].NewRow();
                    newRow["id_LetterLink"] = (int)idLetter;
                    newRow["id_TagLink"] = (int)tags.IDTagLink;
                    tagsOfLetterDS.Tables[0].Rows.Add(newRow);
                }
                tagsOfLetterAdapter.Update(tagsOfLetterDS);
            }
            catch (Exception tagsException)
            {
                tagsErrNo = tagsException.HResult; // Возвращаем код ошибки
            }
            finally
            {
                tagsOfLetterAdapter.Dispose();
            }
            #endregion
            return tagsErrNo;
        }

        /// <summary>
        /// Метод получения списка существующих в БД тэгов
        /// </summary>
        /// <returns> Кортеж
        /// Первый параметр = список тэгов
        /// Второй параметр = код ошибки
        /// </returns>
        public Tuple<List<Tags>, int> GetTags()
        {
            int selectErrNo = 1; // Этот код говорит о том, что ошибки нет (если все пройдет успешно, код не изменится)
            List<Tags> tagsList = new List<Tags>();
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString))
            {
                #region Блок определения SQL команды
                SqlCommand selectCommand = new SqlCommand()
                {
                    Connection = connectionDocsVision,
                    CommandText = selectTagsString,
                    CommandType = CommandType.StoredProcedure
                };
                #endregion
                #region Блок try-catch-finally
                try
                {
                    connectionDocsVision.Open();
                    SqlDataReader selectReader = selectCommand.ExecuteReader();
                    foreach (DbDataRecord dr in selectReader)
                    {
                        tagsList.Add(new Tags()
                        {
                            IDTag = Int32.Parse(dr["id_Tag"].ToString()),
                            TagName = dr["tagName"].ToString()
                        });
                    }
                    connectionDocsVision.Close();
                }
                catch (Exception tagsException)
                {
                    selectErrNo = tagsException.HResult; // Возвращаем код ошибки
                    tagsList = null;                     // Очищаем список
                }
                finally
                {
                    connectionDocsVision.Dispose();
                }
                #endregion
            }
            return new Tuple<List<Tags>, int>(tagsList, selectErrNo);
        }

        /// <summary>
        /// Метод удаления тэга из списка тэгов
        /// Внимание! Удаление тэга, который связан отношение многие-ко-многим с таблицей писем,
        /// приведет к удалению этого тэга из промежуточной таблицы.
        /// Данный тэг станет недоступным для всех писем (или удалится у них)
        /// </summary>
        /// <param name="idTag"> Идентификатор удаляемого тэга </param>
        /// <returns> Код ошибки при удалении </returns>
        public int DeleteTag(int idTag)
        {
            int deleteErrNo = 1; // Этот код говорит о том, что ошибки нет (если все пройдет успешно, код не изменится)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString))
            {
                #region Блок определения SQL команды
                SqlCommand deleteCommand = new SqlCommand()
                {
                    Connection  = connectionDocsVision,
                    CommandText = deleteTagString,
                    CommandType = CommandType.StoredProcedure
                };
                SqlParameter deleteParam = new SqlParameter()
                {
                    ParameterName = "@IDDELETETAG",
                    Value = idTag,
                    Direction = ParameterDirection.Input
                };
                deleteCommand.Parameters.Add(deleteParam);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    connectionDocsVision.Open();
                    deleteCommand.ExecuteNonQuery();
                    connectionDocsVision.Close();
                }
                catch (Exception tagsException)
                {
                    deleteErrNo = tagsException.HResult; // Возвращаем код ошибки
                }
                finally
                {
                    connectionDocsVision.Dispose();
                }
                #endregion
            }
            return deleteErrNo;
        }

        /// <summary>
        /// Метод добавления нового тэга в список доступных тэгов
        /// </summary>
        /// <param name="tag"> Создаваемый тэг </param>
        /// <returns> Код ошибки при добавлении </returns>
        public int AddTag(string tagName)
        {
            int addErrNo = 1; // Этот код говорит о том, что ошибки нет (если все пройдет успешно, код не изменится)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString))
            {
                #region Блок определения SQL команды
                SqlCommand addCommand = new SqlCommand()
                {
                    Connection = connectionDocsVision,
                    CommandText = addTagString,
                    CommandType = CommandType.StoredProcedure
                };
                SqlParameter addParam = new SqlParameter()
                {
                    ParameterName = "@TAGNAME",
                    Value = tagName,
                    Direction = ParameterDirection.Input
                };
                addCommand.Parameters.Add(addParam);
                #endregion
                #region Блок try-catch-finally
                try
                {
                    connectionDocsVision.Open();
                    addCommand.ExecuteNonQuery();
                    connectionDocsVision.Close();
                }
                catch (Exception tagsException)
                {
                    addErrNo = tagsException.HResult; // Возвращаем код ошибки
                }
                finally
                {
                    connectionDocsVision.Dispose();
                }
                #endregion
            }
            return addErrNo;
        }

        /// <summary>
        /// Метод переименования тэга
        /// </summary>
        /// <param name="idTag"> Идентификатор переименовываемого тэга </param>
        /// <param name="tag"> Новое название тэга </param>
        /// <returns> Код ошибки при переименовании </returns>
        public int RenameTag(int idTag, string tagName)
        {
            int renameErrNo = 1; // Этот код говорит о том, что ошибки нет (если все пройдет успешно, код не изменится)
            using (connectionDocsVision = new SqlConnection(connectionDocsVisionString))
            {
                #region Блок определения SQL команды
                //string updateTagsString = $"UPDATE tbTags SET tagName = '{tag}' WHERE id_Tag = '{idTag}'"; // Пока прямым текстом
                SqlCommand renameCommand = new SqlCommand()
                {
                    Connection = connectionDocsVision,
                    CommandText = renameTagString,
                    CommandType = CommandType.StoredProcedure
                };
                SqlParameter renameParam_1 = new SqlParameter()
                {
                    ParameterName = "@IDTAG",
                    Value = idTag,
                    Direction = ParameterDirection.Input
                };
                renameCommand.Parameters.Add(renameParam_1);
                SqlParameter renameParam_2 = new SqlParameter()
                {
                    ParameterName = "@TAGNAME",
                    Value = tagName,
                    Direction = ParameterDirection.Input
                };
                renameCommand.Parameters.Add(renameParam_2);
                #endregion
                #region Блок try=catch-finally
                try
                {
                    connectionDocsVision.Open();
                    renameCommand.ExecuteNonQuery();
                    connectionDocsVision.Close();
                }
                catch (Exception tagsException)
                {
                    renameErrNo = tagsException.HResult; // Возвращаем код ошибки
                }
                finally
                {
                    connectionDocsVision.Dispose();
                }
                #endregion
            }
            return renameErrNo;
        }

        /// <summary>
        /// Метод уничтожения соединения с сервером
        /// На мой вгляд - метод лишний, так как во всех методах есть уничтожение соединения в блоке finally
        /// </summary>
        public void Dispose()
        {
            connectionDocsVision.Dispose();
        }
    }
}
