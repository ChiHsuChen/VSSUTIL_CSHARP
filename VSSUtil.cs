/*
 * Name: VSSUtil
 * Purpose: 包裝Visual SourceSafe介面使用
 * DateTime: 2014/05/07
 * Author: Chi-Hsu Chen
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SourceSafeTypeLib;

namespace VSSINFO
{
    class VSSUtil
    {
        VSSDatabase vssdb = new VSSDatabase();
        private const string _ADD_RESULT_OK = "OK";
        List<string> lsList = new List<string>();
        List<string> lsProj = new List<string>();
        
        // 0. 從指定的scrSafe.ini開啟VSS DB
        // lsPath ex: \\localhost\SPC\srcsafe.ini
        public void OpenVSS(string lsPath, string lsAccount, string lsPasswd)
        {            
            try
            {
                CloseVSSDB();
                vssdb.Open(lsPath, lsAccount, lsPasswd);

                return;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return;
            }
        }

        // 1. 取回檔案最新版本編號
        // lsPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        public int getFileLatestVersionNumber(string lsPath)
        {
            VSSItem vssitem;

            try
            {
                vssitem = vssdb.get_VSSItem(lsPath);
            
                return vssitem.VersionNumber;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return -1;
            }
        }

        // 2. 取回檔案的ItemName
        // lsPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        public string getFileItemName(string lsPath)
        {
            VSSItem vssitem;

            try
            {
                vssitem = vssdb.get_VSSItem(lsPath);
            
                return vssitem.Name;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return string.Empty;
            }
        }

        // 3. 取回檔案最後修改日期
        // lsPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        public string getFileModifyDate(string lsPath)
        {
            string lsFileDateTime=string.Empty;
            VSSItem vssitem;

            try
            {
                vssitem = vssdb.get_VSSItem(lsPath);

                foreach (VSSVersion v in vssitem.get_Versions())
                {
                    lsFileDateTime = v.Date.ToString("yyyy-MM-dd HH:mm:ss");  // 以24小時制顯示
                    break;
                }

                return lsFileDateTime;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return string.Empty;
            }
        }

        // 4. by VersionNumber取回特定檔案內容
        // lsPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        // lsVersionNumber ex: 1
        // lsLocalPath ex: c:\Log.vb
        public void getFileByVersionNumber(string lsPath, int lsVersionNumber, string lsLocalPath)
        {
            VSSItem vssitem;

            try
            {
                vssitem = vssdb.get_VSSItem(lsPath);
                //檢查給定版號是否存在
                if (IsGetVersionNumberExceedLatest(lsPath, lsVersionNumber))
                {
                    //throw new Exception("Version " + lsVersionNumber.ToString() + " Not Found!");
                    System.Windows.Forms.MessageBox.Show("Version " + lsVersionNumber.ToString() + " Not Found!Please confirm this version exist.", "VSS Automation",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }
                vssitem.get_Version(lsVersionNumber.ToString()).Get(ref lsLocalPath, 0);
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return;
            }
        }

        // 5. 取得特定檔案是否為check out
        // lsPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        public bool IsFileCheckOut(string lsPath)
        {
            VSSItem vssitem;

            try
            {
                vssitem = vssdb.get_VSSItem(lsPath);
                return vssitem.IsCheckedOut == 2 ? true : false;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return false;
            }
        }

        //6. 給定版本編號是否超過最新版本編號
        // True為超過；False為沒有超過
        private bool IsGetVersionNumberExceedLatest(string lsPath, int lsVersionNumber)
        {
            if (lsVersionNumber > getFileLatestVersionNumber (lsPath))
                return true;
            else
                return false;
        }

        // 7. 比對特定版本檔案是否與local copy相同
        // lsPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        // lsLocalPath ex: d:\log.vb
        // lsVersionNumber ex: 1
        public string IsFileDiffByVersionNumber(string lsPath, string lsLocalPath, int lsVersionNumber)
        {
            VSSItem vssitem;

            try
            {
                if (System.IO.File.Exists(lsLocalPath) == false) return "本地檔案不存在";
                
                vssitem = vssdb.get_VSSItem(lsPath);
                return vssitem.get_Version(lsVersionNumber.ToString()).get_IsDifferent(lsLocalPath).ToString();
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return "N/A";
            }
        }

        // 8. 抓取最新版本的comment
        // lsPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        public string getFileComment(string lsPath)
        {
            string lsComment=string.Empty;
            VSSItem vssitem;

            try
            {
                vssitem = vssdb.get_VSSItem(lsPath);

                foreach (VSSVersion v in vssitem.get_Versions())
                {
                    lsComment = v.Comment;
                    break;
                }
                return lsComment == string.Empty ? "NULL" : lsComment;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return string.Empty;
            }
        }

        // 9. 顯示錯誤訊息
        private void ShowErrorMessage(System.Runtime.InteropServices.COMException cmo)
        {
            System.Windows.Forms.MessageBox.Show("Execption Msg = \n" + cmo.ToString(), "VSS Automation",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

            return;
        }

        // 10. 顯示目前開啟專案資訊
        public string getProjectInfo()
        {
            return vssdb.CurrentProject;
        }

        // 11. 傳回登入使用者名稱 
        public string getLogOnUser()
        {
            return vssdb.Username;            
        }

        // 12. close vssdb
        public void CloseVSSDB()
        {
            vssdb.Close();

            return;
        }

        // 13. 加入一個檔案到現有專案中
        // lsVssPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        // lsLocalPath ex: d:\log.vb
        // lsComment ex: 加入comment
        public string addFile(string lsVssPath,string lsLocalPath,string lsComment)
        {
            VSSItem vssitem;
            
            try
            {
                vssitem = vssdb.get_VSSItem(lsVssPath);
                vssitem.Add(lsLocalPath, lsComment, 0);
                
                return _ADD_RESULT_OK;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return cmo.ToString();
            }
        }
        
        // 14. 回傳DBNAME
        public string getDBName()
        {
            return vssdb.DatabaseName;
        }

        // 15. 取回最後修改人
        // lsPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        public string getLastUpdateUser(string lsPath)
        {
            string lsLastUpdateUser = string.Empty;
            VSSItem vssitem;

            try
            {
                vssitem = vssdb.get_VSSItem(lsPath);

                foreach (VSSVersion v in vssitem.get_Versions())
                {
                    lsLastUpdateUser = v.Username;
                    break;
                }
                return lsLastUpdateUser == string.Empty ? "NULL" : lsLastUpdateUser;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return string.Empty;
            }
        }
        
        // 16. 建立專案子目錄
        // lsVssPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        // lsSubFolder ex: FORM
        public void createSubFolder(string lsVssProjPath, string lsSubFolder)
        {
            VSSItem vssitem;

            try
            {
                vssitem = vssdb.get_VSSItem(lsVssProjPath);
                vssitem.NewSubproject(lsSubFolder);

                return;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return;
            }
        }

        // 17. 取得目前開啟的srcSafe.ini
        public string getSrcSafe_INI()
        {
            return vssdb.SrcSafeIni;
        }

        // 18. 判斷是否目前的檔案存在
        public bool IsSpecifiedVssItemExist(string lsVssPath, string lsFindPath)
        {
            if (lsVssPath.Trim()==string.Empty)
            {
                lsVssPath = lsFindPath.Substring(0, lsFindPath.LastIndexOf('/') + 1);
            }

            lsList.Clear();
            collectInfoByPath(lsVssPath);
            if (lsList == null) return false;
#if (DEBUG)
            foreach (string s in lsList)
            {
                Console.WriteLine("path=" + s.ToString());
            }
#endif
            
            foreach (string s in lsList)
            {
                if (s.ToLower() == lsFindPath.ToLower())
                    return true;
            }
            return false;
        }

        // 19. 從給定的專案路徑中收集檔案項目
        public List<string> collectInfoByPath(string lsVssPath)
        {
            try
            {
                VSSItem lsProj = vssdb.get_VSSItem(lsVssPath, false);

                foreach (VSSItem item in lsProj.get_Items(false))
                {
                    if (item.Type == (int)VSSItemType.VSSITEM_PROJECT)
                        collectInfoByPath(item.Spec);
                    else
                        lsList.Add(lsVssPath + (lsVssPath.Substring(lsVssPath.Length-1,1) != "/" ? "/" : string.Empty) + item.Name);
                }
               return lsList;            
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                return null; 
            }
        }

        // 20. 從給定的專案路徑中收集專案
        public void collectProjInfoByPath(string lsVssPath)
        {
            try
            {
                VSSItem Proj = vssdb.get_VSSItem(lsVssPath, false);

                if (Proj.get_Items(false).Count == 0) 
                    lsProj.Add(lsVssPath + (lsVssPath.Substring(lsVssPath.Length - 1, 1) != "/" ? "/" : string.Empty));
                
                foreach (VSSItem item in Proj.get_Items(false))
                {
                    if (item.Type == (int)VSSItemType.VSSITEM_PROJECT)
                    {
                        lsProj.Add(lsVssPath + (lsVssPath.Substring(lsVssPath.Length - 1, 1) != "/" ? "/" : string.Empty));
                        collectProjInfoByPath(item.Spec);
                    }
                }
                return;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                return;
            }
        }

        // 21. 指定的VSS Project是否存在
        public bool IsSpecifiedVssProjectExist(string lsVssPath, string lsFindPath)
        {
            if (lsVssPath.Trim() == string.Empty)
            {
                lsVssPath = lsFindPath.Substring(0, lsFindPath.LastIndexOf('/') + 1);
            }

            lsProj.Clear();
            collectProjInfoByPath(lsVssPath);
            if (lsProj == null) return false;
#if (DEBUG)
            foreach (string s in lsProj)
            {
                Console.WriteLine("path=" + s.ToString());
            }
#endif
            foreach (string s in lsProj)
            {
                if (s.ToLower() == lsFindPath.ToLower())
                    return true;
            }
            return false;
        }

        // 22. 以內容比對與特定版號內容是否相同
        // lsVssPath ex: $/SPC/WindowsApplication1.root/WindowsApplication1/WindowsApplication1/Log.vb
        // lsLocalPath ex: d:\log.vb
        public int getVersionNumberByContent(string lsVssPath, string lsLocalPath)
        {
            VSSItem vssitem;

            try
            {
                if (System.IO.File.Exists(lsLocalPath) == false)
                    return 0;
                
                vssitem = vssdb.get_VSSItem(lsVssPath);

                foreach (VSSVersion v in vssitem.get_Versions())
                {
                    // 如果檔案內容一樣，則視為與該版號內容相同，回傳該版號
                    if (IsFileDiffByVersionNumber(lsVssPath, lsLocalPath, v.VersionNumber).ToLower() == "false")
                        return v.VersionNumber;
                }
                return 0;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                return 0;
            }
        }

        public List<string> getProjItemList(string lsFindPath)
        {
            lsProj.Clear();
            collectProjInfoByPath(lsFindPath);
            if (lsProj == null) return null;
#if (DEBUG)
            foreach (string s in lsProj)
            {
                Console.WriteLine("path=" + s.ToString());
            }
#endif

            return lsProj;
        }

        // 尚不開放使用 - 讀取專案下所有項目
        private string readProject(string lsPath)
        {
            string lsContent="";
            try
            {
                VSSItem lsProj = vssdb.get_VSSItem(lsPath, false);
                foreach (VSSItem item in lsProj.get_Items(false))
                {
                    if (item.Type == (int)VSSItemType.VSSITEM_PROJECT)
                    {
                        readProject(item.Spec);
                    }
                    else
                    {
                        string.Concat(lsContent, "<Item=>");
                        string.Concat(lsContent, item.Name);
                        string.Concat(lsContent, "<VersionNumber=>");
                        string.Concat(lsContent, item.VersionNumber.ToString());
                        string.Concat(lsContent, "\n");
                    }
                }
                return lsContent;
            }
            catch (System.Runtime.InteropServices.COMException cmo)
            {
                ShowErrorMessage(cmo);
                return string.Empty;
            }
        }

    }
}
