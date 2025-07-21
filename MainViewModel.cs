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
        [ObservableProperty]
        private string projectPath = Directory.GetCurrentDirectory();

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
        private string itemUnit = "";

        [ObservableProperty]
        private double itemPerPrice = 1;

        [ObservableProperty]
        private double itemNum = 1;

        [ObservableProperty]
        private string itemDesc = "";

        [ObservableProperty]
        private ObservableCollection<GeneralItem> userItems = new ObservableCollection<GeneralItem>();

        [ObservableProperty]
        private GeneralItem? selectedUserItem;

        [ObservableProperty]
        private string versionText = $"CarrotMRO\r\nVer.{Assembly.GetExecutingAssembly().GetName().Version}";


        public MainViewModel()
        {
            // 更新datagrid直接编辑的元素值到autosave
            UserItems.CollectionChanged += (s, e) => {
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
                var autoSaveExcelFilePath = Path.Combine(ProjectPath, "autosave.xlsx");
                ExcelHelper.WriteToExcel(UserItems.ToList(), autoSaveExcelFilePath, true, ExcelItemFactory.AutoSaveHeader, ExcelItemFactory.AutosaveWriteFactory);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        partial void OnSelectedStandardItemNameChanged(string value)
        {
            GeneralItem? matchItem = StandardItems.FirstOrDefault(i => i.Name == SelectedStandardItemName);
            if (matchItem != null)
            {
                ItemPerPrice = matchItem.PerPrice;
                ItemUnit = matchItem.Unit;
                ItemDesc = matchItem.Description;
            }
        }

        [RelayCommand]
        public void OpenProjectFolder()
        {
            try
            {
                OpenFolderDialog openFolderDialog = new OpenFolderDialog() {
                    InitialDirectory = ProjectPath,
                };
                if (openFolderDialog.ShowDialog() == true)
                {
                    ProjectPath = openFolderDialog.FolderName;
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
                var standardExcelFilePath = Path.Combine(ProjectPath, "standard.xlsx");
                if (File.Exists(standardExcelFilePath))
                {
                    var standardItemsFromExcel = ExcelHelper.ReadFromExcel(standardExcelFilePath, ExcelItemFactory.StandardReadFactory);
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
                var autoSaveExcelFilePath = Path.Combine(ProjectPath, "autosave.xlsx");
                if (File.Exists(autoSaveExcelFilePath))
                {
                    var autosaveItemsFromExcel = ExcelHelper.ReadFromExcel(autoSaveExcelFilePath, ExcelItemFactory.AutosaveReadFactory);
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
                if (ItemUnit != matchItem.Unit)
                {
                    MessageBox.Show($"当前单位({ItemUnit})与标准单位({matchItem.Unit})不匹配");
                    return;
                }
                if (ItemPerPrice != matchItem.PerPrice)
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
                UserItems.Add(new GeneralItem() {
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
                    if (item.Unit != matchItem.Unit)
                    {
                        stringBuilder.AppendLine($"项目({item.Name}):当前单位({ItemUnit})与标准单位({matchItem.Unit})不匹配");
                    }
                    if (item.PerPrice != matchItem.PerPrice)
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
            var groupedItems = UserItems.GroupBy(item => item.Part);

            var outExcelFilePath = Path.Combine(ProjectPath, "out.xlsx");
            ExcelHelper.WriteToExcel(groupedItems, outExcelFilePath, true, ExcelItemFactory.OutHeader, ExcelItemFactory.OutWriteFactory);
        }
    }
}