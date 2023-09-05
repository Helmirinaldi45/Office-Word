using Microsoft.Win32; // Required for file dialogs
using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Office.Interop.Word;
using Xabe.FFmpeg;

namespace WpfApp32;

public partial class MainWindow : System.Windows.Window
{
    private string currentFilePath; // Store the current file path
    private ScreenCaptureRecorder recorder;


    public MainWindow()
    {
        InitializeComponent();
    }

    // New
    private void New_Click(object sender, RoutedEventArgs e)
    {
        // Clear the existing content in the RichTextBox
        richTextEditor.Document.Blocks.Clear();
    }

    // New Tab
    private void NewTab_Click(object sender, RoutedEventArgs e)
    {
        // Create a new tab and add a RichTextBox to it
        TabItem newTab = new TabItem();
        newTab.Header = "New Tab";
        RichTextBox richTextBox = new RichTextBox();
        newTab.Content = richTextBox;
        tabControl.Items.Add(newTab);
    }

    // Close Tab
    private void CloseTab_Click(object sender, RoutedEventArgs e)
    {
        if (tabControl.SelectedItem != null)
        {
            tabControl.Items.Remove(tabControl.SelectedItem);
        }
    }

    // Close All Tabs
    private void CloseAllTabs_Click(object sender, RoutedEventArgs e)
    {
        tabControl.Items.Clear();
    }

    // Close
    private void Close_Click(object sender, RoutedEventArgs e)
    {
        // Implement code to close the current document (save changes if necessary)
        if (IsDocumentModified())
        {
            MessageBoxResult result = MessageBox.Show("Do you want to save changes before closing?", "Save Changes", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                Save_Click(sender, e);
            }
            else if (result == MessageBoxResult.Cancel)
            {
                return; // Cancel closing
            }
        }

        // Close the tab or window
        if (tabControl.SelectedItem != null)
        {
            tabControl.Items.Remove(tabControl.SelectedItem);
        }
        else
        {
            this.Close(); // Close the application window
        }
    }

    // Open
    private void Open_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";

        if (openFileDialog.ShowDialog() == true)
        {
            string filePath = openFileDialog.FileName;
            string fileContent = File.ReadAllText(filePath);

            // Create a new tab and load the file content into a RichTextBox
            TabItem newTab = new TabItem();
            newTab.Header = "Opened File: " + Path.GetFileName(filePath);
            RichTextBox richTextBox = new RichTextBox();
            richTextBox.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new Run(fileContent)));
            newTab.Content = richTextBox;
            tabControl.Items.Add(newTab);
        }
    }

    // Save
    private void Save_Click(object sender, RoutedEventArgs e)
    {
        if (currentFilePath != null)
        {
            TextRange textRange = new TextRange(richTextEditor.Document.ContentStart, richTextEditor.Document.ContentEnd);
            File.WriteAllText(currentFilePath, textRange.Text);
        }
        else
        {
            Save_Click(sender, e);
        }
    }

    // Save As Word
    private void SaveAsWord_Click(object sender, RoutedEventArgs e)
    {
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Add();

        // Retrieve the text content from the RichTextBox
        TextRange textRange = new TextRange(richTextEditor.Document.ContentStart, richTextEditor.Document.ContentEnd);
        string text = textRange.Text;

        // Insert the text content into the Word document
        wordDoc.Content.Text = text;

        // Specify the file path for saving the Word document
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "Word Documents (*.docx)|*.docx";
        if (saveFileDialog.ShowDialog() == true)
        {
            string filePath = saveFileDialog.FileName;
            wordDoc.SaveAs2(filePath);
            wordDoc.Close();
            wordApp.Quit();

            MessageBox.Show("Document saved as Word file.", "Save As Word", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
    //event handler for save as pdf
    private void SaveAsPdf_Click(object sender, RoutedEventArgs e)
    {
        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
        Microsoft.Office.Interop.Word.Document wordDoc = wordApp.Documents.Add();

        // Retrieve the text content from the RichTextBox
        TextRange textRange = new TextRange(richTextEditor.Document.ContentStart, richTextEditor.Document.ContentEnd);
        string text = textRange.Text;

        // Insert the text content into the Word document
        wordDoc.Content.Text = text;

        // Specify the file path for saving the Word document
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "PDF Documents (*.pdf)|*.pdf";
        if (saveFileDialog.ShowDialog() == true)
        {
            string filePath = saveFileDialog.FileName;
            wordDoc.SaveAs2(filePath, WdSaveFormat.wdFormatPDF);
            wordDoc.Close();
            wordApp.Quit();

            MessageBox.Show("Document saved as PDF file.", "Save As PDF", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
    private void SaveAsImage_Click(object sender, RoutedEventArgs e)
    {
        // Create a RenderTargetBitmap to capture the RichTextBox content
        RenderTargetBitmap renderTargetBitmap = new RenderTargetBitmap((int)richTextEditor.ActualWidth, (int)richTextEditor.ActualHeight, 96, 96, PixelFormats.Default);
        renderTargetBitmap.Render(richTextEditor);

        // Encode the bitmap as a PNG image
        PngBitmapEncoder pngEncoder = new PngBitmapEncoder();
        pngEncoder.Frames.Add(BitmapFrame.Create(renderTargetBitmap));

        // Specify the file path for saving the image
        SaveFileDialog saveFileDialog = new SaveFileDialog();
        saveFileDialog.Filter = "PNG Images (*.png)|*.png";
        if (saveFileDialog.ShowDialog() == true)
        {
            using (FileStream stream = new FileStream(saveFileDialog.FileName, FileMode.Create))
            {
                pngEncoder.Save(stream);
            }

            MessageBox.Show("Document saved as an image.", "Save As Image", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
    // Exit
    private void Exit_Click(object sender, RoutedEventArgs e)
    {
        // Implement code to exit the application (save changes if necessary)
        if (IsDocumentModified())
        {
            MessageBoxResult result = MessageBox.Show("Do you want to save changes before exiting?", "Save Changes", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                Save_Click(sender, e);
            }
            else if (result == MessageBoxResult.Cancel)
            {
                return; // Cancel exiting
            }
        }

        this.Close(); // Close the application window
    }
    //event handler for undo
    private void Undo_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Undo();
    }
    //event handler for redo
    private void Redo_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Redo();
    }
    //event handler for cut
    private void Cut_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Cut();
    }
    //event handler for copy
    private void Copy_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Copy();
    }
    //event handler for paste
    private void Paste_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Paste();
    }
    //event handler for select all
    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.SelectAll();
    }
    //event handler for font bold
    private void Bold_Click(object sender, RoutedEventArgs e)
    {
        if (richTextEditor.Selection.GetPropertyValue(Run.FontWeightProperty).Equals(FontWeights.Bold))
        {
            richTextEditor.Selection.ApplyPropertyValue(Run.FontWeightProperty, FontWeights.Normal);
        }
        else
        {
            richTextEditor.Selection.ApplyPropertyValue(Run.FontWeightProperty, FontWeights.Bold);
        }
    }
    //event handler for font italic
    private void Italic_Click(object sender, RoutedEventArgs e)
    {
        if (richTextEditor.Selection.GetPropertyValue(Run.FontStyleProperty).Equals(FontStyles.Italic))
        {
            richTextEditor.Selection.ApplyPropertyValue(Run.FontStyleProperty, FontStyles.Normal);
        }
        else
        {
            richTextEditor.Selection.ApplyPropertyValue(Run.FontStyleProperty, FontStyles.Italic);
        }
    }
    //event handler for font underline
    private void Underline_Click(object sender, RoutedEventArgs e)
    {
        if (richTextEditor.Selection.GetPropertyValue(Run.TextDecorationsProperty).Equals(TextDecorations.Underline))
        {
            richTextEditor.Selection.ApplyPropertyValue(Run.TextDecorationsProperty, null);
        }
        else
        {
            richTextEditor.Selection.ApplyPropertyValue(Run.TextDecorationsProperty, TextDecorations.Underline);
        }
    }
    //event handler for background image
    private void BackgroundImage_Click(object sender, RoutedEventArgs e)
    {
        // Create an OpenFileDialog
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "Image Files (*.bmp, *.jpg, *.png)|*.bmp;*.jpg;*.png";

        // Show the dialog and check if the user selected an image
        if (openFileDialog.ShowDialog() == true)
        {
            // Create a new BitmapImage and use SetSource to load the image
            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.UriSource = new Uri(openFileDialog.FileName);
            bitmapImage.EndInit();

            // Create a new Image element and set its Source to the BitmapImage
            Image image = new Image();
            image.Source = bitmapImage;

            // Create a new InlineUIContainer using the Image element
            InlineUIContainer container = new InlineUIContainer(image);

            // Insert the image at the current caret position
            richTextEditor.CaretPosition.Paragraph.Inlines.Add(container);
        }
    } 
    //event handler for strike through
    private void StrikeThrough_Click(object sender, RoutedEventArgs e)
    {
        if (richTextEditor.Selection.GetPropertyValue(Run.TextDecorationsProperty).Equals(TextDecorations.Strikethrough))
        {
            richTextEditor.Selection.ApplyPropertyValue(Run.TextDecorationsProperty, null);
        }
        else
        {
            richTextEditor.Selection.ApplyPropertyValue(Run.TextDecorationsProperty, TextDecorations.Strikethrough);
        }
    }
    //event handler for insert image
    private void InsertImage_Click(object sender, RoutedEventArgs e)
    {
        // Create an OpenFileDialog
        OpenFileDialog openFileDialog = new OpenFileDialog();
        openFileDialog.Filter = "Image Files (*.bmp, *.jpg, *.png)|*.bmp;*.jpg;*.png";

        // Show the dialog and check if the user selected an image
        if (openFileDialog.ShowDialog() == true)
        {
            // Create a new BitmapImage and use SetSource to load the image
            BitmapImage bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.UriSource = new Uri(openFileDialog.FileName);
            bitmapImage.EndInit();

            // Create a new Image element and set its Source to the BitmapImage
            Image image = new Image();
            image.Source = bitmapImage;

            // Create a new InlineUIContainer using the Image element
            InlineUIContainer container = new InlineUIContainer(image);

            // Insert the image at the current caret position
            richTextEditor.CaretPosition.Paragraph.Inlines.Add(container);
        }
    }
    //event handler for insert table
    //event handler for insert hyperlink
    private void InsertHyperlink_Click(object sender, RoutedEventArgs e)
    {
        // Create a new Hyperlink with default text and navigate to the URL
        System.Windows.Documents.Hyperlink hyperlink = new System.Windows.Documents.Hyperlink(new Run("http://www.google.com"));
        hyperlink.NavigateUri = new Uri("http://www.google.com");

        // Insert the hyperlink at the current caret position
        richTextEditor.CaretPosition.Paragraph.Inlines.Add(hyperlink);
    }
    //event handler for layout => letter
    private void Letter_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 816;
        richTextEditor.Document.PageHeight = 1056;
    }
    //event handler for layout => legal
    private void Legal_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 816;
        richTextEditor.Document.PageHeight = 1344;
    }
    //event handler for layout => A3
    private void A3_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 1122;
        richTextEditor.Document.PageHeight = 1587;
    }
    //event handler for layout => A4
    private void A4_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 816;
        richTextEditor.Document.PageHeight = 1122;
    }
    //event handler for layout => A5
    private void A5_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 583;
        richTextEditor.Document.PageHeight = 816;
    }
    //event handler for layout => A6
    private void A6_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 409;
        richTextEditor.Document.PageHeight = 583;
    }
    //event handler for layout => A7
    private void A7_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 291;
        richTextEditor.Document.PageHeight = 409;
    }
    //event handler for layout => A8
    private void A8_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 204;
        richTextEditor.Document.PageHeight = 291;
    }
    //event handler for layout => A9
    private void A9_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 145;
        richTextEditor.Document.PageHeight = 204;
    }
    //event handler for text box selection changed
    //event handler for new page
    private void NewPage_Click(object sender, RoutedEventArgs e)
    {
        System.Windows.Documents.Paragraph paragraph = new System.Windows.Documents.Paragraph();
        richTextEditor.Document.Blocks.Add(paragraph);
    }
    //event handler for new window
    private void NewWindow_Click(object sender, RoutedEventArgs e)
    {
        MainWindow mainWindow = new MainWindow();
        mainWindow.Show();
    }
    //event handler for ledger
    private void Ledger_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageHeight = 1632;
        richTextEditor.Document.PageWidth = 1056;
    }
    //event handler for layout => A10
    private void A10_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 102;
        richTextEditor.Document.PageHeight = 145;
    }
    //event handler for A1
    private void A1_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 2384;
        richTextEditor.Document.PageHeight = 3370;
    }
    //event handler for delete
    private void Delete_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Selection.Text = "";
    }
    //event handler for A2
    private void A2_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 1684;
        richTextEditor.Document.PageHeight = 2384;
    }
    //event handler for A0
    private void A0_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 3370;
        richTextEditor.Document.PageHeight = 4768;
    }
    //event handler for thesis
    private void Thesis_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 816;
        richTextEditor.Document.PageHeight = 1056;
    }
    //event handler for layout => Tabloid
    private void Tabloid_Click(object sender, RoutedEventArgs e)
    {
        richTextEditor.Document.PageWidth = 1056;
        richTextEditor.Document.PageHeight = 1632;
    }
    private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        // Ambil nilai dari Slider
        double sliderValue = slider.Value;

        // Konversi nilai Slider menjadi nilai Hue dalam HSL
        double hue = sliderValue;

        // Ubah warna latar belakang RichTextBox
        SolidColorBrush backgroundBrush = new SolidColorBrush(ColorFromHsl(hue, 1.0, 0.5));
        richTextEditor.Background = backgroundBrush;
    }

    // Fungsi untuk mengubah nilai HSL menjadi warna RGB
    private System.Windows.Media.Color ColorFromHsl(double hue, double saturation, double lightness)
    {
        double chroma = (1 - Math.Abs(2 * lightness - 1)) * saturation;
        double huePrime = hue / 60;
        double x = chroma * (1 - Math.Abs(huePrime % 2 - 1));
        double r, g, b;

        if (0 <= huePrime && huePrime < 1)
        {
            r = chroma;
            g = x;
            b = 0;
        }
        else if (1 <= huePrime && huePrime < 2)
        {
            r = x;
            g = chroma;
            b = 0;
        }
        else if (2 <= huePrime && huePrime < 3)
        {
            r = 0;
            g = chroma;
            b = x;
        }
        else if (3 <= huePrime && huePrime < 4)
        {
            r = 0;
            g = x;
            b = chroma;
        }
        else if (4 <= huePrime && huePrime < 5)
        {
            r = x;
            g = 0;
            b = chroma;
        }
        else
        {
            r = chroma;
            g = 0;
            b = x;
        }

        double m = lightness - chroma / 2;
        byte byteR = (byte)((r + m) * 255);
        byte byteG = (byte)((g + m) * 255);
        byte byteB = (byte)((b + m) * 255);

        return System.Windows.Media.Color.FromRgb(byteR, byteG, byteB);
    }
    private void FontSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (richTextEditor != null && fontSizeSlider != null)
        {
            richTextEditor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, fontSizeSlider.Value);
        }
    }
    private void LineSpacingSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (richTextEditor != null && lineSpacingSlider != null)
        {
            double lineHeight = lineSpacingSlider.Value * fontSizeSlider.Value; // Sesuaikan dengan nilai yang Anda inginkan.
            richTextEditor.Selection.ApplyPropertyValue(System.Windows.Documents.Paragraph.LineHeightProperty, lineHeight);
        }
    }
    private void ColorSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (richTextEditor != null)
        {
            byte red = (byte)redSlider.Value;
            byte green = (byte)greenSlider.Value;
            byte blue = (byte)blueSlider.Value;

            System.Windows.Media.Color textColor = System.Windows.Media.Color.FromRgb(red, green, blue);
            SolidColorBrush brush = new SolidColorBrush(textColor);

            richTextEditor.Selection.ApplyPropertyValue(TextElement.ForegroundProperty, brush);
        }
    }
    private void LayoutSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (richTextEditor != null)
        {
            double newSize = layoutSizeSlider.Value * 1080; // Sesuaikan faktor perubahan ukuran sesuai dengan kebutuhan Anda.
            richTextEditor.Width = newSize; // Sesuaikan dengan properti yang Anda ingin ubah, seperti Width atau Height.dou
            richTextEditor.Height = newSize;
        }
    }
    private void PageSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (richTextEditor != null)
        {
            int desiredPageCount = (int)PageSlider.Value;

            // Pastikan nilai minimum halaman adalah 1.
            if (desiredPageCount < 1)
            {
                desiredPageCount = 2;
                PageSlider.Value = 1;
            }

            // Hitung faktor skala teks untuk mengatur jumlah halaman.
            double scale = 1.0 / desiredPageCount;

            // Hitung margin atas dan bawah untuk mengatur jumlah halaman.
            double topMargin = 100 * scale; // Atur sesuai kebutuhan Anda.
            double bottomMargin = 100 * scale; // Atur sesuai kebutuhan Anda.

            // Terapkan faktor skala teks.
            richTextEditor.Selection.ApplyPropertyValue(TextElement.FontSizeProperty, 12 * scale); // Atur sesuai kebutuhan Anda.

            // Terapkan margin atas dan bawah.
            richTextEditor.Margin = new Thickness(0, topMargin, 0, bottomMargin);

            // Selain itu, Anda dapat melakukan penyesuaian lainnya seperti pengaturan line spacing atau tindakan lain yang sesuai.
        }
    }
    private void StartRecordingButton_Click(object sender, RoutedEventArgs e)
    {
        // Buat instance rekaman dan atur pengaturannya.
        recorder = new ScreenCaptureRecorder();
        recorder.CaptureRectangle = new System.Drawing.Rectangle((int)richTextEditor.PointToScreen(new System.Windows.Point(0, 0)).X,
                                                               (int)richTextEditor.PointToScreen(new System.Windows.Point(0, 0)).Y,
                                                               (int)richTextEditor.ActualWidth,
                                                               (int)richTextEditor.ActualHeight);
        recorder.OutputPath = "output.mp4";
        recorder.VideoCodec = VideoCodec.h261;

        // Mulai merekam.
        recorder.Start();

        // Aktifkan tombol "Hentikan Rekaman" dan nonaktifkan tombol "Mulai Rekaman".
        startRecordingButton.IsEnabled = false;
        stopRecordingButton.IsEnabled = true;
    }
    private void StopRecordingButton_Click(object sender, RoutedEventArgs e)
    {
        if (recorder != null)
        {
            // Hentikan rekaman.
            recorder.Stop();
            recorder.Dispose();

            // Nonaktifkan tombol "Hentikan Rekaman" dan aktifkan tombol "Mulai Rekaman".
            startRecordingButton.IsEnabled = true;
            stopRecordingButton.IsEnabled = false;
        }
    }


    // Helper method to check if the document has been modified
    private bool IsDocumentModified()
    {
        TextRange textRange = new TextRange(richTextEditor.Document.ContentStart, richTextEditor.Document.ContentEnd);
        return !string.Equals(textRange.Text, File.ReadAllText(currentFilePath));
    }
}
