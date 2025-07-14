using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;

namespace CarrotMRO
{
    public partial class MainViewModel : ObservableObject
    {
        [ObservableProperty]
        private string projectPath = Directory.GetCurrentDirectory();

        [ObservableProperty]
        private ObservableCollection<GeneralItem> standardItems = new ObservableCollection<GeneralItem>();

        [ObservableProperty]
        private ObservableCollection<string> parts = new ObservableCollection<string>()
        {
            "<默认类目>",
        };

        [ObservableProperty]
        private string selectedPart = "<默认类目>";

        [ObservableProperty]
        private ObservableCollection<string> customItemNames = new ObservableCollection<string>();

        [ObservableProperty]
        private string selectedCustomItemName = "<项目1>";

        [ObservableProperty]
        private string selectedStandardItemName = "";

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
            var autoSaveExcelFilePath = Path.Combine(ProjectPath, "autosave.xlsx");
            ExcelHelper.WriteToExcel(UserItems.ToList(), autoSaveExcelFilePath, true, ExcelItemFactory.AutoSaveHeader, ExcelItemFactory.AutosaveWriteFactory);
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
        public void AddItem()
        {
            GeneralItem? matchItem = StandardItems.FirstOrDefault(i => i.Name == SelectedStandardItemName);
            if (matchItem != null)
            {
                UserItems.Add(new GeneralItem()
                {
                    Part = SelectedPart,
                    Name = SelectedStandardItemName,
                    CustomName = SelectedCustomItemName,
                    Num = ItemNum,
                    Description = ItemDesc,
                    Unit = matchItem.Unit,
                    PerPrice = matchItem.PerPrice,
                });
            }
            else
            {
                MessageBox.Show("未选择列表中的标准项目");
            }
        }


        [RelayCommand]
        public void ValidateUserItems()
        {
            // TODO
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
    }
}