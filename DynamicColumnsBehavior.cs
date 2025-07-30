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
            if (behavior.AssociatedObject != null) // Ensure DataGrid is attached before updating
            {
                behavior.UpdateColumns();
            }
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
                // Create the base binding object
                var binding = new Binding(columnDef.BindingPath)
                {
                    // Set a sensible default for updates
                    UpdateSourceTrigger = columnDef.UpdateSourceTrigger,
                    Mode = columnDef.BindingMode
                };

                // Check if a format string exists in the definition...
                if (!string.IsNullOrEmpty(columnDef.StringFormat))
                {
                    //    ...and apply it to the binding object.
                    binding.StringFormat = columnDef.StringFormat;
                }

                // Create the column and assign the fully configured binding to it
                var column = new DataGridTextColumn
                {
                    Header = columnDef.Header,
                    Binding = binding
                };

                AssociatedObject.Columns.Add(column);
            }
        }
    }
}