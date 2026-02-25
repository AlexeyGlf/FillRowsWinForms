using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private FillHandleDataGridView dataGridView;
        private Button btnTest;
        private Label lblInstructions;
        private Panel topPanel;
        public Form1()
        {
            InitializeComponent();
            this.Text = "Excel-like Fill Handle - РАБОЧАЯ ВЕРСИЯ";
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // Верхняя панель с инструкциями
            topPanel = new Panel();
            topPanel.Dock = DockStyle.Top;
            topPanel.Height = 80;
            topPanel.BackColor = Color.FromArgb(240, 240, 240);
            topPanel.Padding = new Padding(10);

            // Заголовок
            Label lblTitle = new Label();
            lblTitle.Text = "Маркер заполнения как в Excel";
            lblTitle.Font = new Font("Segoe UI", 14, FontStyle.Bold);
            lblTitle.Location = new Point(10, 10);
            lblTitle.AutoSize = true;

            // Инструкция
            lblInstructions = new Label();
            lblInstructions.Text = "1. Выделите одну или несколько ячеек мышью\n" +
                                   "2. Наведите на синий квадратик в правом нижнем углу выделения\n" +
                                   "3. Когда курсор станет крестиком - тяните вниз или вправо\n" +
                                   "4. Отпустите мышь - ячейки заполнятся";
            lblInstructions.Font = new Font("Segoe UI", 10);
            lblInstructions.Location = new Point(10, 40);
            lblInstructions.AutoSize = true;
            lblInstructions.ForeColor = Color.FromArgb(64, 64, 64);

            topPanel.Controls.Add(lblTitle);
            topPanel.Controls.Add(lblInstructions);

            // Кнопка для сброса данных
            btnTest = new Button();
            btnTest.Text = "Сбросить данные";
            btnTest.Location = new Point(750, 25);
            btnTest.Size = new Size(120, 30);
            btnTest.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnTest.Click += (s, e) => {
                dataGridView.SetSampleData();
                dataGridView.ClearSelection();
            };

            topPanel.Controls.Add(btnTest);

            // Создаем DataGridView с маркером
            dataGridView = new FillHandleDataGridView();
            dataGridView.Dock = DockStyle.Fill;
            dataGridView.SetSampleData();

            // Добавляем контролы на форму
            this.Controls.Add(dataGridView);
            this.Controls.Add(topPanel);
        }
    }
    public class FillHandleDataGridView : DataGridView
    {
        // Свойства для маркера заполнения
        private bool _isFillHandleVisible = false;
        private Rectangle _fillHandleRect;
        private bool _isDragging = false;
        private Point _dragStartCell;
        private Point _dragCurrentCell;
        private Color _fillHandleColor = Color.FromArgb(0, 120, 215); // Синий как в Excel
        private int _fillHandleSize = 10; // Чуть больше для удобства
        private List<DataGridViewCell> _selectedCellsBackup;

        public FillHandleDataGridView()
        {
            // Включаем выделение ячеек
            this.SelectionMode = DataGridViewSelectionMode.CellSelect;
            this.MultiSelect = true;

            // Включаем двойную буферизацию для плавной отрисовки
            this.DoubleBuffered = true;

            // Подписываемся на события
            this.Paint += DataGridView_Paint;
            this.MouseDown += DataGridView_MouseDown;
            this.MouseMove += DataGridView_MouseMove;
            this.MouseUp += DataGridView_MouseUp;

            // Настраиваем внешний вид
            this.BackgroundColor = Color.White;
            this.GridColor = Color.LightGray;
            this.DefaultCellStyle.SelectionBackColor = Color.FromArgb(189, 214, 254);
            this.DefaultCellStyle.SelectionForeColor = Color.Black;
            this.RowHeadersWidth = 50;
            this.ColumnHeadersHeight = 30;
        }

        // Отрисовка маркера и рамки
        private void DataGridView_Paint(object sender, PaintEventArgs e)
        {
            // Всегда проверяем и обновляем позицию маркера
            UpdateFillHandlePosition();

            // Рисуем маркер, если есть выделенные ячейки
            if (_isFillHandleVisible && !_isDragging)
            {
                DrawFillHandle(e.Graphics);
            }

            // Рисуем рамку при перетаскивании
            if (_isDragging)
            {
                DrawDragRectangle(e.Graphics);
            }
        }

        // Обновление позиции маркера
        private void UpdateFillHandlePosition()
        {
            if (this.SelectedCells.Count > 0)
            {
                // Находим нижнюю правую ячейку выделения
                DataGridViewCell lastCell = GetLastSelectedCell();

                if (lastCell != null && lastCell.RowIndex >= 0 && lastCell.ColumnIndex >= 0)
                {
                    Rectangle cellRect = this.GetCellDisplayRectangle(lastCell.ColumnIndex, lastCell.RowIndex, false);

                    _fillHandleRect = new Rectangle(
                        cellRect.Right - _fillHandleSize,
                        cellRect.Bottom - _fillHandleSize,
                        _fillHandleSize,
                        _fillHandleSize
                    );

                    _isFillHandleVisible = true;
                }
            }
            else
            {
                _isFillHandleVisible = false;
            }
        }

        // Получение последней выделенной ячейки (для позиции маркера)
        private DataGridViewCell GetLastSelectedCell()
        {
            DataGridViewCell lastCell = null;
            int maxRow = -1;
            int maxCol = -1;

            foreach (DataGridViewCell cell in this.SelectedCells)
            {
                if (cell.RowIndex >= maxRow && cell.ColumnIndex >= maxCol)
                {
                    maxRow = cell.RowIndex;
                    maxCol = cell.ColumnIndex;
                    lastCell = cell;
                }
            }
            return lastCell;
        }

        // Рисование маркера
        private void DrawFillHandle(Graphics g)
        {
            using (SolidBrush brush = new SolidBrush(_fillHandleColor))
            {
                g.FillRectangle(brush, _fillHandleRect);
            }

            // Рисуем белую обводку для контраста
            using (Pen pen = new Pen(Color.White, 1))
            {
                g.DrawRectangle(pen, _fillHandleRect);
            }
        }

        // Рисование рамки при перетаскивании
        private void DrawDragRectangle(Graphics g)
        {
            if (_dragStartCell != null && _dragCurrentCell != null)
            {
                int minRow = Math.Min(_dragStartCell.Y, _dragCurrentCell.Y);
                int maxRow = Math.Max(_dragStartCell.Y, _dragCurrentCell.Y);
                int minCol = Math.Min(_dragStartCell.X, _dragCurrentCell.X);
                int maxCol = Math.Max(_dragStartCell.X, _dragCurrentCell.X);

                Rectangle dragRect = Rectangle.Empty;
                for (int row = minRow; row <= maxRow; row++)
                {
                    for (int col = minCol; col <= maxCol; col++)
                    {
                        if (row >= 0 && row < this.Rows.Count &&
                            col >= 0 && col < this.Columns.Count)
                        {
                            Rectangle cellRect = this.GetCellDisplayRectangle(col, row, false);
                            if (dragRect.IsEmpty)
                                dragRect = cellRect;
                            else
                                dragRect = Rectangle.Union(dragRect, cellRect);
                        }
                    }
                }

                if (!dragRect.IsEmpty)
                {
                    // Рисуем полупрозрачную заливку
                    using (SolidBrush brush = new SolidBrush(Color.FromArgb(50, _fillHandleColor)))
                    {
                        g.FillRectangle(brush, dragRect);
                    }

                    // Рисуем рамку
                    using (Pen pen = new Pen(_fillHandleColor, 2))
                    {
                        pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                        g.DrawRectangle(pen, dragRect);
                    }
                }
            }
        }

        // Обработка нажатия мыши
        private void DataGridView_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && _isFillHandleVisible && _fillHandleRect.Contains(e.Location))
            {
                // Сохраняем выделенные ячейки
                _selectedCellsBackup = new List<DataGridViewCell>();
                foreach (DataGridViewCell cell in this.SelectedCells)
                {
                    _selectedCellsBackup.Add(cell);
                }

                // Начинаем перетаскивание
                _isDragging = true;

                var hit = this.HitTest(e.X, e.Y);
                if (hit.Type == DataGridViewHitTestType.Cell)
                {
                    _dragStartCell = new Point(hit.ColumnIndex, hit.RowIndex);
                    _dragCurrentCell = _dragStartCell;
                }

                this.Cursor = Cursors.Cross;
                this.Invalidate();
                //e.Handled = true;
            }
        }

        // Обработка движения мыши
        private void DataGridView_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDragging)
            {
                var hit = this.HitTest(e.X, e.Y);
                if (hit.Type == DataGridViewHitTestType.Cell)
                {
                    if (_dragCurrentCell.X != hit.ColumnIndex || _dragCurrentCell.Y != hit.RowIndex)
                    {
                        _dragCurrentCell = new Point(hit.ColumnIndex, hit.RowIndex);
                        this.Invalidate(); // Перерисовываем для обновления рамки
                    }
                }
            }
            else
            {
                // Меняем курсор при наведении на маркер
                if (_isFillHandleVisible && _fillHandleRect.Contains(e.Location))
                {
                    this.Cursor = Cursors.Cross;
                }
                else
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        // Обработка отпускания мыши
        private void DataGridView_MouseUp(object sender, MouseEventArgs e)
        {
            if (_isDragging && e.Button == MouseButtons.Left)
            {
                // Заполняем ячейки
                FillCells();

                _isDragging = false;
                _isFillHandleVisible = false;
                this.Invalidate();
                this.Cursor = Cursors.Default;
            }
        }

        // Заполнение ячеек
        private void FillCells()
        {
            if (_dragStartCell == null || _dragCurrentCell == null || _selectedCellsBackup == null)
                return;

            int minRow = Math.Min(_dragStartCell.Y, _dragCurrentCell.Y);
            int maxRow = Math.Max(_dragStartCell.Y, _dragCurrentCell.Y);
            int minCol = Math.Min(_dragStartCell.X, _dragCurrentCell.X);
            int maxCol = Math.Max(_dragStartCell.X, _dragCurrentCell.X);

            // Начинаем транзакцию обновления
            this.BeginInvoke(new MethodInvoker(() =>
            {
                try
                {
                    // Заполняем все ячейки в области
                    for (int row = minRow; row <= maxRow; row++)
                    {
                        for (int col = minCol; col <= maxCol; col++)
                        {
                            // Проверяем, не является ли ячейка исходной
                            bool isSourceCell = false;
                            foreach (DataGridViewCell sourceCell in _selectedCellsBackup)
                            {
                                if (sourceCell.RowIndex == row && sourceCell.ColumnIndex == col)
                                {
                                    isSourceCell = true;
                                    break;
                                }
                            }

                            if (!isSourceCell && row < this.Rows.Count && col < this.Columns.Count)
                            {
                                // Копируем значение из первой выделенной ячейки
                                if (_selectedCellsBackup.Count > 0)
                                {
                                    object valueToCopy = _selectedCellsBackup[0].Value;

                                    // Если копируем несколько ячеек, пытаемся создать ряд
                                    if (_selectedCellsBackup.Count > 1)
                                    {
                                        // Пробуем распознать числовой ряд
                                        if (IsNumericSeries(_selectedCellsBackup, out List<object> series))
                                        {
                                            int index = (row - minRow) * (maxCol - minCol + 1) + (col - minCol);
                                            if (index < series.Count)
                                            {
                                                valueToCopy = series[index];
                                            }
                                        }
                                    }

                                    this.Rows[row].Cells[col].Value = valueToCopy;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при заполнении: {ex.Message}");
                }
            }));
        }

        // Проверка на числовой ряд
        private bool IsNumericSeries(List<DataGridViewCell> cells, out List<object> series)
        {
            series = new List<object>();

            // Пробуем распарсить все значения как числа
            List<double> numbers = new List<double>();
            foreach (var cell in cells)
            {
                if (double.TryParse(cell.Value?.ToString(), out double num))
                {
                    numbers.Add(num);
                }
                else
                {
                    return false;
                }
            }

            // Проверяем арифметическую прогрессию
            if (numbers.Count >= 2)
            {
                double step = numbers[1] - numbers[0];
                bool isArithmetic = true;

                for (int i = 1; i < numbers.Count; i++)
                {
                    if (Math.Abs(numbers[i] - (numbers[0] + step * i)) > 0.0001)
                    {
                        isArithmetic = false;
                        break;
                    }
                }

                if (isArithmetic)
                {
                    // Генерируем продолжение ряда
                    for (int i = 0; i < 100; i++) // Максимум 100 ячеек
                    {
                        series.Add(numbers[0] + step * i);
                    }
                    return true;
                }
            }

            return false;
        }

        // Пример данных для тестирования
        public void SetSampleData()
        {
            this.ColumnCount = 7;
            this.RowCount = 15;

            for (int i = 0; i < 7; i++)
            {
                this.Columns[i].Name = $"Колонка {i + 1}";
                this.Columns[i].Width = 90;
                this.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            // Заполняем тестовыми данными
            Random rand = new Random();
            for (int row = 0; row < 15; row++)
            {
                for (int col = 0; col < 7; col++)
                {
                    if (col == 0)
                        this.Rows[row].Cells[col].Value = row + 1; // Числа для теста ряда
                    else if (col == 1)
                        this.Rows[row].Cells[col].Value = $"Текст {row + 1}";
                    else
                        this.Rows[row].Cells[col].Value = rand.Next(1, 100);
                }
            }
        }
    }
}

