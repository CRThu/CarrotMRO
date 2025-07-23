using CarrotMRO;
using Microsoft.Xaml.Behaviors;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace CarrotMRO
{
    public class DynamicColumnsBehavior : Behavior<DataGrid>
    {
        public static readonly DependencyProperty ColumnsProperty =
            DependencyProperty.Register(
                "Columns",
                typeof(ObservableCollection<DataGridColumnDefinition>),
                typeof(DynamicColumnsBehavior),
                new PropertyMetadata(null, OnColumnsChanged));

        public ObservableCollection<DataGridColumnDefinition> Columns
        {
            get => (ObservableCollection<DataGridColumnDefinition>)GetValue(ColumnsProperty);
            set => SetValue(ColumnsProperty, value);
        }

        private static void OnColumnsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var behavior = (DynamicColumnsBehavior)d;
            behavior.UpdateColumns();
        }

        protected override void OnAttached()
        {
            base.OnAttached();
            AssociatedObject.AutoGenerateColumns = false;
            UpdateColumns();
        }

        private void UpdateColumns()
        {
            if (AssociatedObject == null || Columns == null) return;

            AssociatedObject.Columns.Clear();
            foreach (var columnDef in Columns)
            {
                var column = new DataGridTextColumn {
                    Header = columnDef.Header,
                    Binding = new Binding(columnDef.BindingPath) {
                        Mode = columnDef.BindingMode,
                        UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                    }
                };
                AssociatedObject.Columns.Add(column);
            }
        }
    }
}