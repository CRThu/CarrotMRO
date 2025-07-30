using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Windows;

namespace CarrotMRO
{
    public partial class MainViewModel : ObservableObject
    {
        private AppConfig _appConfig;

        [ObservableProperty]
        private string projectDirectory = Directory.GetCurrentDirectory();

        [ObservableProperty]
        private ObservableCollection<DataGridColumnDefinition> standardDataGridColumns;

        [ObservableProperty]
        private ObservableCollection<DataGridColumnDefinition> quotationDataGridColumns;


        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(FilteredStandardItemNames))]
        private ObservableCollection<GeneralItem> standardItems = new ObservableCollection<GeneralItem>();

        [ObservableProperty]
        private ObservableCollection<string> parts = new ObservableCollection<string>()
        {
            "<默认类目>",
        };

        [ObservableProperty]
        private string selectedPart = "<默认类目>";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(FilteredCustomItemNames))]
        private ObservableCollection<string> customItemNames = new ObservableCollection<string>();

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(FilteredCustomItemNames))]
        private string selectedCustomItemName = "<项目1>";

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(FilteredStandardItemNames))]
        private string selectedStandardItemName = "";

        public IEnumerable<string> FilteredCustomItemNames => CustomItemNames.Where(item => item.Contains(SelectedCustomItemName, StringComparison.CurrentCultureIgnoreCase));

        public IEnumerable<string> FilteredStandardItemNames => StandardItems.Select(i => i.Name).Where(item => item.Contains(SelectedStandardItemName, StringComparison.CurrentCultureIgnoreCase));

        [ObservableProperty]
        private GeneralItem newItem;

        [ObservableProperty]
        private ObservableCollection<GeneralItem> userItems = new ObservableCollection<GeneralItem>();

        [ObservableProperty]
        private GeneralItem? selectedUserItem;

        [ObservableProperty]
        private string versionText = $"CarrotMRO\r\nVer.{Assembly.GetExecutingAssembly().GetName().Version}";


        public MainViewModel()
        {
            // 更新datagrid直接编辑的元素值到autosave
            UserItems.CollectionChanged += (s, e) =>
            {
                if (e.NewItems != null && e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Add)
                {
                    foreach (GeneralItem item in e.NewItems)
                        item.PropertyChanged += (s, e) => CurrentItemsAutosave();
                }

                if (e.OldItems != null && (
                        e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Remove
                        || e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Replace
                        || e.Action == System.Collections.Specialized.NotifyCollectionChangedAction.Reset))
                {
                    foreach (GeneralItem item in e.OldItems)
                        item.PropertyChanged -= (s, e) => CurrentItemsAutosave();
                }

                CurrentItemsAutosave();
            };
        }

        private void CurrentItemsAutosave()
        {
            try
            {
                var autoSaveExcelFilePath = Path.Combine(ProjectDirectory, _appConfig.Autosave.FileName);
                ExcelHelper.WriteToExcel(_appConfig.Autosave.Header, UserItems.ToList(), autoSaveExcelFilePath, true, ExcelItemFactory.ItemWrite);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        partial void OnSelectedPartChanged(string value)
        {
            NewItem.Part = SelectedPart;
        }

        partial void OnSelectedCustomItemNameChanged(string value)
        {
            NewItem.CustomName = SelectedCustomItemName;
        }

        partial void OnSelectedStandardItemNameChanged(string value)
        {
            NewItem.Name = SelectedStandardItemName;

            GeneralItem? matchItem = StandardItems.FirstOrDefault(i => i.Name == SelectedStandardItemName);
            if (matchItem != null)
            {
                if (_appConfig.Match.Unit)
                    NewItem.Unit = matchItem.Unit;
                if (_appConfig.Match.Num)
                    NewItem.Num = matchItem.Num;
                if (_appConfig.Match.PerPrice)
                {
                    NewItem.BaseMaterialInPerPrice = matchItem.BaseMaterialInPerPrice;
                    NewItem.AuxMaterialInPerPrice = matchItem.AuxMaterialInPerPrice;
                    NewItem.MaterialInPerPrice = matchItem.MaterialInPerPrice;
                    NewItem.MachineInPerPrice = matchItem.MachineInPerPrice;
                    NewItem.LaborInPerPrice = matchItem.LaborInPerPrice;
                    NewItem.PerPrice = matchItem.PerPrice;
                }
                if (_appConfig.Match.Desc)
                    NewItem.Description = matchItem.Description;
            }
        }

        [RelayCommand]
        public void BrowseProject()
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog()
                {
                    InitialDirectory = ProjectDirectory,
                    Filter = "YAML 文件|*.yaml;*.yml|所有文件|*.*",
                    DefaultExt = "yaml"
                };
                if (ofd.ShowDialog() == true)
                {
                    ProjectDirectory = Path.GetDirectoryName(ofd.FileName)!;

                    // 读取工程配置
                    _appConfig = ConfigLoader.LoadFromFile(ofd.FileName);

                    // 更新DataGrid列
                    StandardDataGridColumns = new ObservableCollection<DataGridColumnDefinition>(_appConfig.Standard.Header.BuildColumns());
                    QuotationDataGridColumns = new ObservableCollection<DataGridColumnDefinition>(_appConfig.Quotation.Header.BuildColumns());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        [RelayCommand]
        public void LoadProject()
        {
            try
            {
                // 读取标准模板
                var standardExcelFilePath = Path.Combine(ProjectDirectory, _appConfig.Standard.FileName);
                if (File.Exists(standardExcelFilePath))
                {
                    var standardItemsFromExcel = ExcelHelper.ReadFromExcel(_appConfig.Standard.Header, standardExcelFilePath, ExcelItemFactory.ItemRead);
                    StandardItems.Clear();
                    for (int i = 0; i < standardItemsFromExcel.Count; i++)
                    {
                        StandardItems.Add(standardItemsFromExcel[i]);
                    }
                }
                else
                {
                    MessageBox.Show($"未找到标准文件({standardExcelFilePath})");
                }

                // 读取软件自动保存记录
                var autoSaveExcelFilePath = Path.Combine(ProjectDirectory, _appConfig.Autosave.FileName);
                if (File.Exists(autoSaveExcelFilePath))
                {
                    var autosaveItemsFromExcel = ExcelHelper.ReadFromExcel(_appConfig.Autosave.Header, autoSaveExcelFilePath, ExcelItemFactory.ItemRead);
                    UserItems.Clear();
                    for (int i = 0; i < autosaveItemsFromExcel.Count; i++)
                    {
                        UserItems.Add(autosaveItemsFromExcel[i]);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        [RelayCommand]
        public void ValidateNewItem()
        {
            try
            {
                GeneralItem? matchItem = StandardItems.FirstOrDefault(i => i.Name == SelectedStandardItemName);
                if (matchItem == null)
                {
                    MessageBox.Show("未匹配到标准项目");
                    return;
                }
                if (_appConfig.Validate.Unit && ItemUnit != matchItem.Unit)
                {
                    MessageBox.Show($"当前单位({ItemUnit})与标准单位({matchItem.Unit})不匹配");
                    return;
                }
                if (_appConfig.Validate.PerPrice && ItemPerPrice != matchItem.PerPrice)
                {
                    MessageBox.Show($"当前单价({ItemPerPrice})与标准单价({matchItem.PerPrice})不匹配");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        [RelayCommand]
        public void AddNewItem()
        {
            try
            {
                UserItems.Add(new GeneralItem()
                {
                    Part = SelectedPart,
                    Name = SelectedStandardItemName,
                    CustomName = SelectedCustomItemName,
                    Num = ItemNum,
                    Description = ItemDesc,
                    Unit = ItemUnit,
                    PerPrice = ItemPerPrice,
                });

                if (!CustomItemNames.Contains(SelectedCustomItemName))
                    CustomItemNames.Add(SelectedCustomItemName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }


        [RelayCommand]
        public void ValidateUserItems()
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (var item in UserItems)
            {
                GeneralItem? matchItem = StandardItems.FirstOrDefault(i => i.Name == item.Name);
                if (matchItem == null)
                {
                    stringBuilder.AppendLine($"项目({item.Name}):未匹配到标准项目");
                }
                else
                {
                    if (_appConfig.Validate.Unit && item.Unit != matchItem.Unit)
                    {
                        stringBuilder.AppendLine($"项目({item.Name}):当前单位({ItemUnit})与标准单位({matchItem.Unit})不匹配");
                    }
                    if (_appConfig.Validate.PerPrice && item.PerPrice != matchItem.PerPrice)
                    {
                        stringBuilder.AppendLine($"项目({item.Name}):当前单价({ItemPerPrice})与标准单价({matchItem.PerPrice})不匹配");
                    }
                }
            }

            if (stringBuilder.Length > 0)
            {
                MessageBox.Show(stringBuilder.ToString());
            }
            else
            {
                MessageBox.Show("验证未发现问题");
            }
        }

        [RelayCommand]
        public void RemoveSelectedUserItem()
        {
            if (SelectedUserItem != null)
            {
                UserItems.Remove(SelectedUserItem);
            }
        }

        [RelayCommand]
        public void ClearAllUserItems()
        {
            if (MessageBox.Show("确认删除所有条目?", "警告", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                UserItems.Clear();
        }

        [RelayCommand]
        public void GenerateReport()
        {
            SaveFileDialog sfd = new SaveFileDialog()
            {
                Title = "保存Excel文件",
                Filter = "Excel文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*",
                DefaultExt = ".xlsx",
                FileName = $"{DateTime.Now:yyyyMMdd}-报价单.xlsx",
                InitialDirectory = ProjectDirectory,
                OverwritePrompt = true
            };

            if (sfd.ShowDialog() == true)
            {
                if (_appConfig.Quotation.Group)
                {
                    var groupedItems = UserItems.GroupBy(item => item.Part);
                    ExcelHelper.WriteToExcel(_appConfig.Quotation.Header, groupedItems, sfd.FileName, true, ExcelItemFactory.ItemWrite);
                }
                else
                {
                    MessageBox.Show("尚未支持");
                }
            }
        }
    }
}