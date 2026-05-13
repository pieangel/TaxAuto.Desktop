using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;

namespace TaxAuto.Desktop.Models
{
    public class OcrJobItem : INotifyPropertyChanged
    {
        public string FilePath { get; }

        public string FileName => Path.GetFileName(FilePath);

        private OcrJobStatus _status = OcrJobStatus.Waiting;
        public OcrJobStatus Status
        {
            get => _status;
            set
            {
                _status = value;
                OnPropertyChanged();
            }
        }

        private TimeSpan? _elapsed;
        public TimeSpan? Elapsed
        {
            get => _elapsed;
            set
            {
                _elapsed = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ElapsedText));
            }
        }

        public string ElapsedText =>
            Elapsed == null ? "" : $"{Elapsed.Value.TotalSeconds:0.0}초";

        private string? _resultJsonPath;
        public string? ResultJsonPath
        {
            get => _resultJsonPath;
            set
            {
                _resultJsonPath = value;
                OnPropertyChanged();
            }
        }

        private string _message = "";
        public string Message
        {
            get => _message;
            set
            {
                _message = value;
                OnPropertyChanged();
            }
        }

        public OcrJobItem(string filePath)
        {
            FilePath = filePath;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
