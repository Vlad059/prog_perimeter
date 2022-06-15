namespace prog_perimeter;


public partial class Form1 : Form
{
    //Массив путей файлов
    public string[] files = new string[] {};

    //Обработка файлов
    public void files_proceesing()
    {
        if(files.Length>0)
        {
            Corel.Interop.CorelDRAW.Application cdr = new Corel.Interop.CorelDRAW.Application();
            cdr.Visible = false;
            //Открытие только 1 файла 
            cdr.OpenDocument(files[0]);
            //Модель corel draw: document - pages - layers - shapes
            var doc_shapes = cdr.ActiveDocument.Pages[0].ActiveLayer.Shapes;
            //Подсчет длины всех shapes
            double shapes_length = 0;                      
            foreach(Corel.Interop.CorelDRAW.Shape shape in doc_shapes)
            {
                shapes_length += shape.Curve.Length;
            }
            cdr.ActiveDocument.Close(); 
            //Коэфициент масштабирования
            //double koef = 25.4;
            MessageBox.Show("length: "+Convert.ToString(shapes_length));                    
        }      
    }

    //Обработчик нажатия на кнопку, который открывает filedialog
    public void on_click_file(object sender, EventArgs e)
    {
        OpenFileDialog open_fd = new OpenFileDialog();
        open_fd.Multiselect = true;
        if (open_fd.ShowDialog() == DialogResult.OK)
        {
            files = open_fd.FileNames;
            files_proceesing(); 
        }           
    }
    //Добавление элементов на форму
    public void AddElements(){
        Label label_message = new Label();
        label_message.Text = "Форма загрузилась";
        label_message.Location = new Point(10,10);
        label_message.AutoSize = true;

        Button but_file_dialog = new Button();
        but_file_dialog.Text = "Выберете файлы";
        but_file_dialog.Location = new Point(10,40);
        but_file_dialog.AutoSize = true;
        but_file_dialog.Click += on_click_file;

        Controls.Add(label_message);
        Controls.Add(but_file_dialog);
    }
    public Form1()
    {
        InitializeComponent();
        AddElements();
             
    }
}
