using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace SsmsResultsGrid.Core.Mvvm
{
    /// <summary>
    /// Single-flight async ICommand: disabled while its task runs, so double-clicks
    /// can't overlap executions. Exceptions route to the injected error handler
    /// instead of crashing the dispatcher.
    /// </summary>
    public sealed class AsyncRelayCommand : ICommand
    {
        private readonly Func<Task> _execute;
        private readonly Func<bool> _canExecute;
        private readonly Action<Exception> _onError;
        private bool _isRunning;

        public AsyncRelayCommand(Func<Task> execute, Func<bool> canExecute = null, Action<Exception> onError = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
            _onError = onError;
        }

        public event EventHandler CanExecuteChanged;

        public bool IsRunning => _isRunning;

        public bool CanExecute(object parameter) => !_isRunning && (_canExecute?.Invoke() ?? true);

        public async void Execute(object parameter)
        {
            if (!CanExecute(parameter)) return;
            _isRunning = true;
            RaiseCanExecuteChanged();
            try
            {
                await _execute().ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                // Cancellation is a normal outcome for refresh/export commands.
            }
            catch (Exception ex)
            {
                _onError?.Invoke(ex);
            }
            finally
            {
                _isRunning = false;
                RaiseCanExecuteChanged();
            }
        }

        public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
    }
}
