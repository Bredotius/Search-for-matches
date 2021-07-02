//Объявление using System дает возможность ссылаться на классы, которые находятся в пространстве имен System,
//так что их можно использовать, не добавляя System. перед именем типа.
using System;
//Интерфейсы и классы, определяющие универсальные коллекции,
//которые позволяют пользователям создавать строго типизированные коллекции.
using System.Collections.Generic;
//Классы, которые могут использоваться в качестве коллекций
//в объектной модели многократно используемой библиотеки.
using System.Collections.ObjectModel;
//Классы, используемые для реализации поведения компонентов
//и элементов управления во время разработки и выполнения.
using System.ComponentModel;
//классы, позволяющие осуществлять взаимодействие с системными процессами,
//журналами событий и счетчиками производительности.
using System.Diagnostics;
//Содержит типы, позволяющие осуществлять чтение и запись в файлы и потоки данных,
//а также типы для базовой поддержки файлов и папок.
using System.IO;
//Содержит классы и интерфейсы, которые поддерживают LINQ (класс Enumerable).
using System.Linq;
//функции средствам записи компилятора, которые используют управляемый код для задания
//в метаданных атрибутов, влияющих на поведение среды CLR во время выполнения.
using System.Runtime.CompilerServices;
//Содержит классы, обеспечивающие доступ к обработчику регулярных выражений платформы .NET Framework. 
using System.Text.RegularExpressions;
//Предоставляет типы, которые упрощают работу по написанию параллельного и асинхронного кода.
using System.Threading.Tasks;
//Предоставляет несколько важных классов базовых элементов Windows Presentation Foundation (WPF),
using System.Windows;
//Предоставляет классы для создания элементов, называемых элементами управления, которые позволяют
//пользователю взаимодействовать с приложением.
using System.Windows.Controls;
//Представляет базовый класс для строгого типизированных классов документов Open XML.
using DocumentFormat.OpenXml.Packaging;
//Представляет класс для работы с диалоговыми окнами.
using Microsoft.WindowsAPICodePack.Dialogs;
//Представляет приложение Word.
using Word = Microsoft.Office.Interop.Word;
//Объявление собственного пространства имен
namespace SearchForMatches
{
    // Создаем класс который отвечает за отображение данных в главном окне.
    // Взаимодействие с главным окном реализуем в классе интерфейс INotifyPropertyChanged.
    public class DataList : INotifyPropertyChanged
    {
        // Поля хранящие значения публичных свойств.
        private string path;
        private string current;
        private int progress;
        private int inFolder;
        private int maximum;
        private int filesCount;
        private int matсhesCount;
        // Свойсва класса.
        public string Path
        {
            get { return path; }
            set
            {
                path = value;
                OnPropertyChanged("Path");
            }
        }
        public string Current
        {
            get { return current; }
            set
            {
                current = value;
                OnPropertyChanged("Current");
            }
        }
        public int Progress
        {
            get { return progress; }
            set
            {
                progress = value;
                OnPropertyChanged("Progress");
            }
        }
        public int InFolder
        {
            get { return inFolder; }
            set
            {
                inFolder = value;
                OnPropertyChanged("InFolder");
            }
        }
        public int Maximum
        {
            get { return maximum; }
            set
            {
                maximum = value;
                OnPropertyChanged("Maximum");
            }
        }
        public int FilesCount
        {
            get { return filesCount; }
            set
            {
                filesCount = value;
                OnPropertyChanged("FilesCount");
            }
        }
        public int MatchesCount
        {
            get { return matсhesCount; }
            set
            {
                matсhesCount = value;
                OnPropertyChanged("MatchesCount");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        // Событие PropertyChanged извещает систему об изменении свойства.
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
    // Класс MainWindow отвечает за логику для окна приложения.
    public partial class MainWindow : Window
    {
        // Функция MainWindow() отвечает компеляцию главного окна и привязку данных к его элементам.
        public MainWindow()
        {
            // Загружаем скомпилированную страницу.
            InitializeComponent();
            // Привязка к элементу главного окна динамической коллекции данных, для отображения результатов работы программы.
            MatchesList.ItemsSource = AllMatches;
        }
        // Шаблон для записи совпадений.
        public class Matches
        {
            // Переменная названия документа с совпадением.
            public string Title { get; set; }
            // Переменная расположения документа.
            public string Path { get; set; }
            // Переменная количества совпадений в документе.
            public int Counter { get; set; }
        }
        // AllMatches представляет динамическую коллекцию совпадений, которая выдает уведомления при добавлении
        // и удалении элементов, а также при обновлении списка.
        public ObservableCollection<Matches> AllMatches { get; set; } = new ObservableCollection<Matches>();
        // dialog хранит в себе информацию о выбраный папке.
        static CommonOpenFileDialog dialog;
        // dialogResult хранит в себе результат взаимоджействия с диалоговым окном выбора файла.
        static CommonFileDialogResult dialogResult;
        // Переменная запрещающая нажимать на кнопку до конца выполнения функции запущеной прошлым нажатием.
        bool click = true;
        // Обработчик для кнопки отвечающей за открытие диалогового окна выбора файла.
        private void SelectFolder_Click(object sender, RoutedEventArgs e)
        {
            if(click == true)
            {
                // Экземпляр класса DataList, который отвечает за информацию о состояние поиска документов.
                DataList info = (DataList)Resources["SearchInfo"];
                // Новыое диалоговое окно.
                dialog = new CommonOpenFileDialog();
                // Делает возможным открывать только папки.
                dialog.IsFolderPicker = true;
                // Отображает стандартное диалоговое окно, которое предлагает пользователю открыть файл.
                dialogResult = dialog.ShowDialog();
                // Условие выполняется если в результате откытия диалогового окна была выбрана папка.
                if (dialogResult == CommonFileDialogResult.Ok)
                {
                    // Свойству Path переменной info присваивается полный путь до папки.
                    info.Path = dialog.FileName;
                    // Скрываются элементы интерфейса отвечающие за отображение результата поиска.
                    ProgressInfo.Visibility = Visibility.Hidden;
                    CheckedInfo.Visibility = Visibility.Hidden;
                }
            }
        }
        // Асинхронный обработчик для кнопки отвечающей за поиск совпадений в документах.
        private async void FindMatches_Click(object sender, RoutedEventArgs e)
        {
            // Условие выполняется если в результате откытия диалогового окна была выбрана папка и не запущена функция, которая активируется этой кнопкой.
            if (dialogResult == CommonFileDialogResult.Ok && click == true)
            {
                // Переменная запрещает нажимать на кнопку.
                click = false;
                // Скрываются элементы интерфейса отвечающие за отображение обработанных документов.
                MatchesList.Visibility = Visibility.Hidden;
                // Открываются элементы интерфейса отвечающие за отображение прогресса поиска.
                ProgressGrid.Visibility = Visibility.Visible;
                // Скрываются элементы интерфейса отвечающие за отображение результата поиска.
                ProgressInfo.Visibility = Visibility.Hidden;
                CheckedInfo.Visibility = Visibility.Hidden;
                // Очищаем результат прошлого поиска.
                AllMatches.Clear();
                // Экземпляр класса DataList, который отвечает за информацию о состояние поиска документов.
                DataList info = (DataList)Resources["SearchInfo"];
                // Обнуление переменных отвечающих за информацию о состояние поиска документов.
                // Общее количество найденых совпадений.
                info.MatchesCount = 0;
                // Количество файлов с совпадениями.
                info.FilesCount = 0;
                // Проверяемый документ.
                info.Current = "Поиск документов";
                // Прогресс проверки документов.
                info.Progress = 0;
                // Получаем ключевое слово введенное пользователем в главном окне.
                string key = GetKey();
                // await остановит выполнение кода, чтобы в этот момент могли обновиться элементы главного.
                await Task.Delay(50);
                // Получаем перечислитель путей всех документов в папке, а также во всех вложенных папках.
                IEnumerable<String> List = GetFiles(dialog.FileName);
                // Свойству InFolder переменной info присваивается общее число найденых документов.
                info.InFolder = List.Count();
                // Свойство Maximum  переменной info используется для отображение прогресса поиска совпадений в файлах.
                info.Maximum = List.Count() - 1;
                // Открываются элементы интерфейса отвечающие за отображение результата поиска.
                ProgressInfo.Visibility = Visibility.Visible;
                CheckedInfo.Visibility = Visibility.Visible;
                // Цикл поиска совпадений в обнаруженных документах.
                foreach (var item in List)
                {
                    // Получаем название файла отделяя его от полного пути до файла.
                    int position = item.LastIndexOf("\\");
                    string title = item.Substring(position + 1);
                    info.Current = title;
                    // await остановит выполнение кода, чтобы в этот момент могли обновиться элементы главного.
                    await Task.Delay(50);
                    // Из документа извлекается текст.
                    string text = GetText(item);
                    // В тексте считается количество совпадений с ключевым словом.
                    int counter = CountMatches(text, key);
                    // await остановит выполнение кода, чтобы в этот момент могли обновиться элементы главного.
                    await Task.Delay(50);
                    // Условие выполняется если в тексте есть 1 или более совпадение.
                    if (counter != 0)
                    {
                        // Добавляем подходящий документ в коллекцию.
                        AllMatches.Add(new Matches
                        {
                            Title = title,
                            Path = item,
                            Counter = counter,
                        });
                        // К общему количеству совпадений пробавляется количество совпадений в данном документе.
                        info.MatchesCount += counter;
                        // Общему количеству файлов с совпадениями добывляется 1.
                        info.FilesCount++;
                    }
                    // К прогрессу проверки документов добывляется 1.
                    info.Progress++;
                }
                // Условие выполняется если документов не обнаружено.
                if (List.Count() == 0)
                {
                    // Отображает пользователю информацию о выполнение программы.
                    info.Current = "Директория не содержит документов Word!";
                    // Скрываются элементы интерфейса отвечающие за отображение результата поиска.
                    ProgressInfo.Visibility = Visibility.Hidden;
                    CheckedInfo.Visibility = Visibility.Hidden;
                }
                // Условие выполняется если совпадений не найдено.
                else if (info.FilesCount == 0)
                {
                    // Отображает пользователю информацию о выполнение программы.
                    info.Current = "Совпадений не найдено!";
                }
                else
                {
                    // Сортируется полученый спосок документов.
                    AllMatches = SortMatches(AllMatches);
                    // Скрываются элементы интерфейса отображающий прогресс поиска.
                    ProgressGrid.Visibility = Visibility.Hidden;
                    // Отображаются документы сооветстующие критериям поиска.
                    MatchesList.Visibility = Visibility.Visible;
                }
                // Разрешается выбрать другую папку или изменить ключевое слово.
                click = true;
            }
        }
        // Определение обработчика изменения выбранного документа, отвечающий за открытие папки в которой он находится.
        private void MatchesList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Условие выполняется если изменением является выбор элемента, а отмена выбора.
            if (MatchesList.SelectedIndex != -1)
            {
                // Получаем значение выбранного элемента.
                Matches selected = (Matches)MatchesList.SelectedItem;
                // Открываем пакпу с выделением документа.
                Process.Start(new ProcessStartInfo("explorer.exe", " /select, " + selected.Path));
            }
            // Снятие выделения с документа.
            MatchesList.SelectedIndex = -1;
        }
        // Функция получения ключевого слова.
        private string GetKey()
        {
            // Определение переменной ключа.
            string key = " ";
            // Условие выполняется если поле ввода ключевого слова заполнено, иначе ключем будет пробел.
            if (KeyWord.Text != "")
            {
                // Ключу присваивается значение поля ввода, в строчном регистре.
                key = KeyWord.Text.ToLower();
            }
            // Возвращение переменной ключа.
            return key;
        }
        // Функция пересчета совпадений в тексте
        private int CountMatches(string text, string key)
        {
            // Определение переменной количества совпадений.
            int count = 0;
            // Переменной позиции присваивается индекс первого найденого совпадения в тексте.
            int pos = text.IndexOf(key, 0);
            // Условие выполняется если документ не пустой.
            if (text.Length > 0)
            {
                // Условие выполняется пока не дойдет до конца документа.
                while (pos != -1)
                {
                    // Переменная позиции заменяется на индекс следующего найденого совпадения.
                    pos = text.IndexOf(key, pos + 1);
                    // Совпадение увеличивается на 1.
                    count++;
                }
            }
            // Возвращение переменной количества совпадений.
            return count;
        }
        // Функция извлечения текста из документаю
        private string GetText(string path)
        {
            // Позиция точки перед форматом файла.
            int position = path.LastIndexOf(".");
            // Получаем значение формата файла.
            string extension = path.Substring(position + 1);
            // Переменная извлеченного текста.
            string parText = "";
            try
            {
                // Если документ имеет расширение doc.
                if (extension == "doc")
                {
                    // Путь к файлу.
                    object fileName = path;
                    // Значения по умолчанию для отсутствующих значений.
                    object missing = System.Reflection.Missing.Value;
                    // Разрешение только на чтение.
                    object readOnly = true;
                    // Представление приложения Microsoft Word.
                    Word.Application app = new Word.Application();
                    // Открывает указанный документ и добавляет его в коллекцию Documents. Возвращает объект Document.
                    app.Documents.Open(ref fileName, ref missing, ref readOnly, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    // Получаем открытый документ.
                    Word.Document doc = app.ActiveDocument;
                    // Извлекаем текст из документа.
                    parText = doc.Content.Text;
                    // Закрываем приложение Microsoft Word.
                    app.Quit();
                }
                // Если документ имеет расширение docx.
                else
                {
                    // Создаем новый экземпляр класса WordprocessingDocument из указанного файла.
                    WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(path, false);
                    // Извлекаем текст из документа.
                    parText = wordprocessingDocument.MainDocumentPart.Document.Body.InnerText;
                }
            }
            catch (Exception ex)
            {
                // Ошибка чтения файла
                Debug.Print(ex.Message);
            }
            // Возвращаем извлеченный текст в нижнем регистре.
            return parText.ToLower();
        }
        // функция сортировки совпадений в найденых документов по убыванию. 
        private ObservableCollection<Matches> SortMatches(ObservableCollection<Matches> collection)
        {
            // Переменная позиции максимального найденого количества совпадений.
            int max;
            // Переменная текущего сравниваемого документа.
            var temp = new ObservableCollection<Matches> { new Matches { Title = "", Path = "", Counter = 0, } };
            // Цикл поиска максимального количества совпадений.
            for (int i = 0; i < collection.Count - 1; i++)
            {
                // Присваиваем максимальному текущию позицию документа.
                max = i;
                // Цикл поочередного сравнения максимального с остальными элементами.
                for (int j = i + 1; j < collection.Count; j++)
                {
                    if (collection[j].Counter > collection[max].Counter)
                    {
                        // Если j-й элемент больше максимального, максимальному приваеваем позицию j-го.
                        max = j;
                    }
                }
                // Меняем местами максимальный и i-й элементы.
                temp[0] = collection[max];
                collection[max] = collection[i];
                collection[i] = temp[0];
            }
            // Возвращение отсортированной коллекции.
            return collection;
        }
        // Функция поиска всех документов в папке, а также во всех вложенных папках.
        private IEnumerable<String> GetFiles(String directory)
        {
            // Создаем список
            var list = new List<String>();
            // Итератор по файлам
            using (var iterator = Directory.EnumerateFiles(directory).GetEnumerator())
            {
                try
                {
                    // Выполняется пока не дойдет до конца перечисляемой колекции.
                    while (iterator.MoveNext())
                    {
                        // Позиция "\" перед именем файла.
                        var pos = iterator.Current.LastIndexOf("\\");
                        // Получаем значение именем файла.
                        var title = iterator.Current.Substring(pos + 1);
                        // Регулярное выражение для поиска документов формата doc и docx.
                        Regex reg = new Regex(@"^([^~][^\$]).*(\.docx?)$");
                        // Если имя файла соответствует регулярному выражению.
                        if (reg.IsMatch(title))
                        {
                            // Добавляем в список
                            list.Add(iterator.Current);
                        }
                    }
                }
                catch (ArgumentException ex)
                {
                    // Путь является строкой нулевой длины, содержит только пробелы или содержит недопустимые символы.
                    // Параметр поиска не является допустимым значением.
                    Debug.Print(ex.Message);
                }
                catch (DirectoryNotFoundException ex)
                {
                    // Путь недопустим, например, ссылается на неподключенный диск.
                    Debug.Print(ex.Message);
                }
                catch (IOException ex)
                {
                    // Путь является именем файла.
                    // Длина указанного пути, имени файла или обоих параметров превышает установленный системой предел.
                    Debug.Print(ex.Message);
                }
                catch (UnauthorizedAccessException ex)
                {
                    // Исключение, возникающее в случае запрета доступа операционной системой 
                    // из-за ошибки ввода-вывода или особого типа ошибки безопасности.
                    Debug.Print(ex.Message);
                }
            }
            // Итератор по директориям
            using (var iterator = Directory.EnumerateDirectories(directory).GetEnumerator())
            {
                while (iterator.MoveNext())
                {
                    try
                    {
                        list.AddRange(GetFiles(iterator.Current));
                    }
                    catch (ArgumentException ex)
                    {
                        // Путь является строкой нулевой длины, содержит только пробелы или содержит недопустимые символы.
                        // Параметр поиска не является допустимым значением.
                        Debug.Print(ex.Message);
                    }
                    catch (DirectoryNotFoundException ex)
                    {
                        // Путь недопустим, например, ссылается на неподключенный диск.
                        Debug.Print(ex.Message);
                    }
                    catch (IOException ex)
                    {
                        // Путь является именем файла.
                        // Длина указанного пути, имени файла или обоих параметров превышает установленный системой предел.
                        Debug.Print(ex.Message);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        // Исключение, возникающее в случае запрета доступа операционной системой 
                        // из-за ошибки ввода-вывода или особого типа ошибки безопасности.
                        Debug.Print(ex.Message);
                    }
                }
            }
            // Возвращаем результат
            return list; 
        }
    }
}
