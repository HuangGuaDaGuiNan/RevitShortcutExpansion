using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System;
using System.Windows.Interop;
using System.Windows.Media;
using System.IO;
using System.Xml;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using System.Collections.ObjectModel;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using UIFramework;
using UIFrameworkServices;

namespace RevitShortcutExpansion
{
    [Transaction(TransactionMode.Manual)]
    class ShortcutTemplate : IExternalCommand
    {
        private readonly string modifyKeyboardShortcutsFile = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Autodesk\\Revit\\" + UtilityService.getProductReleaseName() + "\\KeyboardShortcuts.xml";
        private readonly string directoryFullName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Autodesk\\Revit\\" + UtilityService.getProductReleaseName() + "\\Four";
        private readonly string configFilePath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\Autodesk\\Revit\\" + UtilityService.getProductReleaseName() + "\\Four\\config.txt";
        private Window mainWindow;
        private StackPanel extStackPanel;
        private UIFramework.DataGrid shortcutGrid;
        private Button assignButton;
        private Button removeButton;
        private Button inportButton;
        private Button okButton;
        private Button cancelButton;
        private Dictionary<string, ShortcutItem> changedShortcuts = new Dictionary<string, ShortcutItem>();
        private System.Windows.Controls.ComboBox comboBox = new System.Windows.Controls.ComboBox()
        {
            Name = "SelectShortcutsTemplate",
            FontSize = 11,
            Height = 25,
            Margin = new Thickness(3),
            Width = 100,
            VerticalContentAlignment = VerticalAlignment.Center,
        };
        private SelItemType currentSelItemType;
        private SelItemType previousSelItemType;
        private string currentSelItemName;
        private string initialSelItemname;
        private string previousSelItemName;
        private ObservableCollection<ComboBoxItem> comboBoxItemList = new ObservableCollection<ComboBoxItem>();
        private ObservableCollection<ShortcutItem> shortcutList;
        private readonly string defTemplateItemName = "default";
        private readonly string createTemplateItemName = "create";
        private readonly string newTemplateKeyName = "自定义方案";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ShortcutWindow shortcutWindow = new ShortcutWindow();
            mainWindow = shortcutWindow;
            new WindowInteropHelper(shortcutWindow).Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle;
            shortcutWindow.Loaded += ShortcutWindow_Loaded;
            shortcutWindow.ShowDialog();
            
            return Result.Succeeded;
        }

        private void ShortcutWindow_Loaded(object sender, RoutedEventArgs e)
        {
            //获取关键对象
            DependencyObject shortcutWindow = sender as DependencyObject;
            extStackPanel = GetElementByUid(shortcutWindow, "StackPanel_2") as StackPanel;
            shortcutGrid = GetElementByUid(shortcutWindow, "mShortcutGrid") as UIFramework.DataGrid;
            assignButton = GetElementByName(shortcutWindow, "mAssignButton") as Button;
            removeButton = GetElementByName(shortcutWindow, "mRemoveButton") as Button;
            inportButton = GetElementByUid(shortcutWindow, "mImportButton") as Button;
            okButton = GetElementByUid(shortcutWindow, "mOkButton") as Button;
            cancelButton = GetElementByUid(shortcutWindow, "mCancelButton") as Button;
            shortcutList = shortcutGrid.ItemsSource as ObservableCollection<ShortcutItem>;

            //修改界面
            comboBox.ItemsSource = comboBoxItemList;
            extStackPanel.Children.Add(comboBox);
            LoadComboBoxItemList();

            //绑定事件
            comboBox.SelectionChanged += ComboBox_SelectionChanged;
            assignButton.Click += AssignButton_Click;
            removeButton.Click += RemoveButton_Click;
            inportButton.Click += InportButton_Click;
            okButton.Click += OkButton_Click;
            cancelButton.Click += CancelButton_Click;

            //设置选项
            RefreshSelectItem();

            //记录状态
            currentSelItemType = GetSelItemType();
            previousSelItemType = currentSelItemType;
            currentSelItemName = GetSelItemName();
            previousSelItemName = currentSelItemName;
            initialSelItemname = currentSelItemName;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            currentSelItemName = GetSelItemName();
            currentSelItemType = GetSelItemType();

            //保存修改
            if (currentSelItemName != previousSelItemName && previousSelItemType == SelItemType.Custom)
            {
                SaveShortcutFile(shortcutList, previousSelItemName);
            }

            //更新表单
            if (currentSelItemType == SelItemType.Create)
            {
                InputWindow inputWindow = new InputWindow(GetSuggestedName());
                inputWindow.Owner = mainWindow;
                if (inputWindow.ShowDialog() == true)
                {
                    CreateNewShortcutTemplate(shortcutList, inputWindow.SaveName);
                    LogSelectedItemName(inputWindow.SaveName);
                }
                RefreshSelectItem();
            }
            if(currentSelItemType == SelItemType.Custom || currentSelItemType == SelItemType.Official)
            {
                Dictionary<string, ShortcutItem> dictionary;
                if (currentSelItemType == SelItemType.Official)
                {
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(Properties.Resource.shortcuts_Revit2019);
                    dictionary = ShortcutsLoad(doc);
                }
                else
                {
                    dictionary = ShortcutsLoad(directoryFullName + "\\" + currentSelItemName + ".xml");
                }
                foreach (ShortcutItem shortcutItem in shortcutList)
                {
                    //初始化
                    if (shortcutItem.IsReserved)
                    {
                        continue;
                    }
                    if (shortcutItem.Shortcuts.Count != 0)
                    {
                        shortcutItem.Shortcuts.Clear();
                        changedShortcuts[shortcutItem.CommandId] = new ShortcutItem(shortcutItem);
                    }
                    if (!dictionary.TryGetValue(shortcutItem.CommandId, out ShortcutItem value))
                    {
                        continue;
                    }
                    //判断
                    bool flag = false;
                    foreach (string shortcut in value.Shortcuts)
                    {
                        if (!string.IsNullOrEmpty(shortcut) && !shortcutItem.Shortcuts.Contains(shortcut))
                        {
                            shortcutItem.Shortcuts.Add(shortcut);
                            flag = true;
                        }
                    }
                    if (flag)
                    {
                        changedShortcuts[shortcutItem.CommandId] = new ShortcutItem(shortcutItem);
                    }
                }
            }
            //SaveShortcutChanges();



            previousSelItemType = currentSelItemType;
            previousSelItemName = currentSelItemName;
        }

        private void AssignButton_Click(object sender, RoutedEventArgs e)
        {
            if (GetSelItemType() == SelItemType.Official)
            {
                InputWindow inputWindow = new InputWindow(GetSuggestedName());
                inputWindow.Owner = mainWindow;
                if (inputWindow.ShowDialog() == true)
                {
                    CreateNewShortcutTemplate(shortcutList, inputWindow.SaveName);
                    LogSelectedItemName(inputWindow.SaveName);
                }
                RefreshSelectItem();
            }
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            if (GetSelItemType() == SelItemType.Official)
            {
                InputWindow inputWindow = new InputWindow(GetSuggestedName());
                inputWindow.Owner = mainWindow;
                if (inputWindow.ShowDialog() == true)
                {
                    CreateNewShortcutTemplate(shortcutList, inputWindow.SaveName);
                    LogSelectedItemName(inputWindow.SaveName);
                }
                RefreshSelectItem();
            }
        }

        private void InportButton_Click(object sender, RoutedEventArgs e)
        {
            if (GetSelItemType() == SelItemType.Official)
            {
                InputWindow inputWindow = new InputWindow(GetSuggestedName());
                inputWindow.Owner = mainWindow;
                if (inputWindow.ShowDialog() == true)
                {
                    CreateNewShortcutTemplate(shortcutList, inputWindow.SaveName);
                    LogSelectedItemName(inputWindow.SaveName);
                }
                RefreshSelectItem();
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            currentSelItemName = GetSelItemName();
            currentSelItemType = GetSelItemType();
            //提交更改
            SaveShortcutChanges();
            //更新文件
            if (currentSelItemType == SelItemType.Custom)
            {
                SaveShortcutFile(shortcutList, currentSelItemName);
            }
            //记录当前选择项
            LogSelectedItemName(currentSelItemName);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            currentSelItemName = GetSelItemName();
            currentSelItemType = GetSelItemType();
            
            LogSelectedItemName(initialSelItemname);
            RefreshSelectItem();
            if(new DirectoryInfo(directoryFullName).GetFiles("*.xml", SearchOption.TopDirectoryOnly).Length > 0)
            {
                SaveShortcutChanges();
            }
        }

        private void LoadComboBoxItemList()
        {
            Directory.CreateDirectory(directoryFullName);
            DirectoryInfo directoryInfo = new DirectoryInfo(directoryFullName);
            //默认方案
            ComboBoxItem def_comboBoxItem = new ComboBoxItem
            {
                Name = defTemplateItemName,
                Content = "默认快捷键",
                ToolTip = "Revit默认的快捷键方案"
            };
            comboBoxItemList.Add(def_comboBoxItem);
            //自定义方案
            foreach (FileInfo fileInfo in directoryInfo.GetFiles("*.xml", SearchOption.TopDirectoryOnly))
            {
                comboBoxItemList.Add(CreateNewItem(fileInfo.Name.Replace(".xml", "")));
            }
            //创建方案
            ComboBoxItem create_comboBoxItem = new ComboBoxItem
            {
                Name = createTemplateItemName,
                Content = "<创建快捷键方案>",
                ToolTip = "创建新的快捷键方案"
            };
            comboBoxItemList.Add(create_comboBoxItem);
            //初次使用时，如已修改默认方案，则使用当前方案创建自定义方案
            if (File.Exists(modifyKeyboardShortcutsFile) && directoryInfo.GetFiles("*.xml", SearchOption.TopDirectoryOnly).Length == 0)
            {
                string newName = GetSuggestedName();
                CreateNewShortcutTemplate(shortcutList, newName);
                LogSelectedItemName(newName);
            }
        }

        private ComboBoxItem CreateNewItem(string itemName)
        {
            string shoctcutFilePath = directoryFullName + "\\" + itemName + ".xml";
            if (!File.Exists(shoctcutFilePath)) return null;
            FileInfo fileInfo = new FileInfo(shoctcutFilePath);
            StackPanel stackPanel = new StackPanel();
            TextBlock text = new TextBlock
            {
                Text = "自定义的快捷方案"
            };
            stackPanel.Children.Add(text);
            Border border = new Border
            {
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(204, 204, 204)),
                BorderThickness = new Thickness(0, 1, 0, 0),
                Margin = new Thickness(0, 8, 0, 8)
            };
            stackPanel.Children.Add(border);
            TextBlock text1 = new TextBlock
            {
                Text = "按 F1 键获得更多帮助"
            };
            stackPanel.Children.Add(text1);
            ComboBoxItem comboBoxItem = new ComboBoxItem
            {
                Content = fileInfo.Name.Replace(".xml", ""),
                Tag = fileInfo.FullName,
                ToolTip = stackPanel
            };
            return comboBoxItem;
        }

        private void CreateNewShortcutTemplate(ICollection<ShortcutItem> itemList, string name)
        {
            //创建xml文件
            SaveShortcutFile(itemList, name);
            //创建ui
            comboBoxItemList.Insert(comboBoxItemList.Count - 1, CreateNewItem(name));
        }

        private void SaveShortcutFile(ICollection<ShortcutItem> shortcutList, string name)
        {
            string fileName = directoryFullName + "\\" + name + ".xml";
            if (shortcutList == null || shortcutList.Count == 0 || string.IsNullOrEmpty(fileName))
            {
                return;
            }
            Directory.CreateDirectory(directoryFullName);
            XmlDocument xmlDocument = new XmlDocument();
            XmlNode xmlNode = xmlDocument.CreateElement("Shortcuts");
            xmlDocument.AppendChild(xmlNode);
            foreach (ShortcutItem shortcutItem in shortcutList)
            {
                if (shortcutItem.CommandId != null && !shortcutItem.IsReserved)
                {
                    XmlNode xmlNode1 = xmlDocument.CreateElement("ShortcutItem");
                    WriteAttribute(xmlNode1, "CommandName", shortcutItem.CommandName);
                    WriteAttribute(xmlNode1, "CommandId", shortcutItem.CommandId);
                    WriteAttribute(xmlNode1, "Shortcuts", shortcutItem.ShortcutsRep);
                    WriteAttribute(xmlNode1, "Paths", shortcutItem.CommandPaths);
                    xmlNode.AppendChild(xmlNode1);
                }
            }
            try
            {
                using (Stream outStream = new FileStream(fileName, FileMode.Create))
                {
                    xmlDocument.Save(outStream);
                }
            }
            catch (Exception)
            {
                //to do
            }

            void WriteAttribute(XmlNode node, string attributeName, object obj)
            {
                if (obj != null)
                {
                    XmlAttribute xmlAttribute = node.Attributes.GetNamedItem(attributeName) as XmlAttribute;
                    if (xmlAttribute == null)
                    {
                        xmlAttribute = xmlDocument.CreateAttribute(attributeName);
                        node.Attributes.Append(xmlAttribute);
                    }
                    xmlAttribute.Value = obj.ToString();
                }
            }

        }

        private Dictionary<string, ShortcutItem> ShortcutsLoad(string fileName)
        {
            if (!File.Exists(fileName))
            {
                return null;
            }

            XmlDocument xmlDocument;
            Dictionary<string, ShortcutItem> dictionary = new Dictionary<string, ShortcutItem>();

            try
            {
                using (Stream inStream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    xmlDocument = new XmlDocument();
                    xmlDocument.Load(inStream);
                }
            }
            catch (Exception)
            {
                return null;
            }
            XmlNode rootNode = xmlDocument.DocumentElement;
            if (rootNode == null || !rootNode.HasChildNodes || rootNode.Name != "Shortcuts")
            {
                return null;
            }
            foreach (XmlNode childNode in rootNode.ChildNodes)
            {
                if (childNode.Name != "ShortcutItem" || childNode.Attributes == null)
                {
                    continue;
                }
                string commandId = GetAttributeValue(childNode, "CommandId");
                if (string.IsNullOrEmpty(commandId))
                {
                    continue;
                }
                string shortcuts = GetAttributeValue(childNode, "Shortcuts");
                if (string.IsNullOrEmpty(shortcuts))
                {
                    continue;
                }
                if (!dictionary.TryGetValue(commandId, out ShortcutItem shortcutItem))
                {
                    shortcutItem = new ShortcutItem();
                    shortcutItem.CommandId = commandId;
                    string[] array = shortcuts.Split('#');
                    string[] array2 = array;
                    foreach (string text in array2)
                    {
                        if (!string.IsNullOrEmpty(text))
                        {
                            if (ShortcutKeyManager.IsLegalShortcutKey(text))
                            {
                                shortcutItem.Shortcuts.Add(text);
                            }
                        }
                    }
                    if (shortcutItem.Shortcuts.Count > 0)
                    {
                        dictionary[commandId] = shortcutItem;
                    }
                    continue;
                }
            }

            if (dictionary.Count == 0) return null;

            return dictionary;
        }

        private Dictionary<string, ShortcutItem> ShortcutsLoad(XmlDocument xmlDocument)
        {
            Dictionary<string, ShortcutItem> dictionary = new Dictionary<string, ShortcutItem>();
            XmlNode rootNode = xmlDocument.DocumentElement;
            if (rootNode == null || !rootNode.HasChildNodes || rootNode.Name != "Shortcuts")
            {
                return null;
            }
            foreach (XmlNode childNode in rootNode.ChildNodes)
            {
                if (childNode.Name != "ShortcutItem" || childNode.Attributes == null)
                {
                    continue;
                }
                string commandId = GetAttributeValue(childNode, "CommandId");
                if (string.IsNullOrEmpty(commandId))
                {
                    continue;
                }
                string shortcuts = GetAttributeValue(childNode, "Shortcuts");
                if (string.IsNullOrEmpty(shortcuts))
                {
                    continue;
                }
                if (!dictionary.TryGetValue(commandId, out ShortcutItem shortcutItem))
                {
                    shortcutItem = new ShortcutItem();
                    shortcutItem.CommandId = commandId;
                    string[] array = shortcuts.Split('#');
                    string[] array2 = array;
                    foreach (string text in array2)
                    {
                        if (!string.IsNullOrEmpty(text))
                        {
                            if (ShortcutKeyManager.IsLegalShortcutKey(text))
                            {
                                shortcutItem.Shortcuts.Add(text);
                            }
                        }
                    }
                    if (shortcutItem.Shortcuts.Count > 0)
                    {
                        dictionary[commandId] = shortcutItem;
                    }
                    continue;
                }
            }

            if (dictionary.Count == 0) return null;

            return dictionary;
        }

        private void SaveShortcutChanges()
        {
            if (changedShortcuts.Count == 0)
            {
                return;
            }
            foreach (ShortcutItem shortcutItem in changedShortcuts.Values)
            {
                ShortcutsHelper.Commands[shortcutItem.CommandId].ShortcutsRep = shortcutItem.ShortcutsRep;
            }
            ShortcutsHelper.SaveShortcuts(ShortcutsHelper.Commands.Values);
            KeyboardShortcutService.applyShortcutChanges(changedShortcuts);
            changedShortcuts.Clear();
        }

        private string GetAttributeValue(XmlNode node, string name)
        {
            return (node.Attributes.GetNamedItem(name) as XmlAttribute)?.Value;
        }

        private FrameworkElement GetElementByUid(DependencyObject rootObject,string uid)
        {
            int count = VisualTreeHelper.GetChildrenCount(rootObject);
            for(int i = 0; i < count; i++)
            {
                FrameworkElement element = VisualTreeHelper.GetChild(rootObject, i) as FrameworkElement;
                if (element != null)
                {
                    if(element.Uid == uid) { return element; }
                    FrameworkElement element1 = GetElementByUid(element, uid);
                    if (element1 != null) return element1;
                }
            }
            return null;
        }

        private FrameworkElement GetElementByName(DependencyObject rootObject, string name)
        {
            int count = VisualTreeHelper.GetChildrenCount(rootObject);
            for (int i = 0; i < count; i++)
            {
                FrameworkElement element = VisualTreeHelper.GetChild(rootObject, i) as FrameworkElement;
                if (element != null)
                {
                    if (element.Name == name) { return element; }
                    FrameworkElement element1 = GetElementByName(element, name);
                    if (element1 != null) return element1;
                }
            }
            return null;
        }

        private SelItemType GetSelItemType()
        {
            if (comboBox.SelectedItem is ComboBoxItem comboBoxItem)
            {
                if (comboBoxItem.Name == defTemplateItemName)
                {
                    return SelItemType.Official;
                }
                if (comboBoxItem.Name == createTemplateItemName)
                {
                    return SelItemType.Create;
                }
                return SelItemType.Custom;
            }
            return SelItemType.Error;
        }

        private string GetSelItemName()
        {
            if (comboBox.SelectedItem is ComboBoxItem comboBoxItem)
            {
                return comboBoxItem.Content.ToString();
            }
            return "";
        }

        private string GetSuggestedName()
        {
            int maxNum = 0;
            foreach (ComboBoxItem comboBoxItem in comboBoxItemList)
            {
                string itemContent = comboBoxItem.Content.ToString();
                if (itemContent.StartsWith(newTemplateKeyName))
                {
                    string[] textArray = itemContent.Split(' ');
                    if (textArray.Length > 0 && new Regex("^([1-9][0-9]*)$").IsMatch(textArray.Last()))
                    {
                        int num = int.Parse(textArray.Last());
                        if (num > maxNum) maxNum = num;
                    }
                }
            }
            return newTemplateKeyName + " " + (maxNum + 1).ToString();
        }

        private void RefreshSelectItem()
        {
            try
            {
                using (StreamReader sr = new StreamReader(File.OpenRead(configFilePath), Encoding.Default))
                {
                    string text = sr.ReadLine();
                    foreach (ComboBoxItem item in comboBox.Items)
                    {
                        if (item.Content.ToString() == text)
                        {
                            comboBox.SelectedItem = item;
                            break;
                        }
                    }
                }
            }
            catch (Exception)
            {
                comboBox.SelectedIndex = 0;
            }
        }

        private void LogSelectedItemName(string itemName)
        {
            using (FileStream fs = new FileStream(configFilePath, FileMode.Create))
            {
                StreamWriter sw = new StreamWriter(fs, Encoding.Default);
                sw.Write(itemName);
                sw.Close();
            }
            if (File.Exists(modifyKeyboardShortcutsFile)) File.Delete(modifyKeyboardShortcutsFile);
            string filePath = directoryFullName + "\\" + itemName + ".xml";
            if (File.Exists(filePath)) File.Copy(filePath, modifyKeyboardShortcutsFile);

        }

        private enum SelItemType
        {
            Error,
            Official,
            Custom,
            Create
        };
    }
}
