using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
//using System.Data.SqlClient;
using System.Net;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.Threading;
using System.Diagnostics;
using MySql.Data.MySqlClient;


namespace SearchFromCNKI
{
    public struct Paper
    {       
        public string paperid;        
        public string authors;       
        public string orgs;
    }
    public struct Author
    {
        public string name;
        public string org;
        public int authorid;
       
    }
    public struct AuthorSearch
    {
        public Author author;
        public int number;

    }
    public struct PaperAuthor
    {
        public int authorid;
        public string authorname;
        public string org;
        public string paperid;
    }


    public partial class Form1 : Form
    {
         
        public Form1()
        {
            InitializeComponent();
        }

        List<Author> searchneeds = new List<Author>();
        List<AuthorSearch> searchneedsaved = new List<AuthorSearch>();
        Dictionary<string, int> dicauthornumber = new Dictionary<string, int>();

        bool loading = true;
        int state = 0;
        

        private MySqlConnection connectdb()
        {
            MySqlConnection m_Connection=null;
            string connStr = String.Format("server={0};user id={1}; password={2}; port={3}; database=cnki; pooling=false; charset=utf8",

                              "166.111.7.152","root", "FtrjT6SM7h", 3306);
            try
            {
                m_Connection = new MySqlConnection(connStr);
            }
            catch (MySqlException ex)  
            {
 
                MessageBox.Show( "Error connecting to the server: " + ex.Message);
 
            }
            return m_Connection;
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
             //Application.DoEvents();
            if (webBrowser1.ReadyState != WebBrowserReadyState.Complete)
                return;
            if (searchneeds.Count < 1)
                state = 1;

            switch (state)
            { 
                
                case 0:
                    string searchname = searchneeds.First().name;
                    string searchorg = searchneeds.First().org;
                    if (dicauthornumber.ContainsKey(searchname + searchorg))
                    {                        
                        AuthorSearch authorseachhade = new AuthorSearch();
                        authorseachhade.author = searchneeds.First();
                        authorseachhade.number = dicauthornumber[searchname + searchorg];
                        searchneedsaved.Add(authorseachhade);
                        searchneeds.Remove(searchneeds.First());
                        webBrowser1_DocumentCompleted(null, null);
                    }
                    searchforName_Enp(searchname, searchorg);
                    break;
                case 1:
                    loading = false;
                    break;
                case 2:
                    int number=getPapers();
                    AuthorSearch authorseach=new AuthorSearch();
                    authorseach.author=searchneeds.First();
                    authorseach.number=number;
                    searchneedsaved.Add(authorseach);
                    searchneeds.Remove(searchneeds.First());
                    dicauthornumber[authorseach.author.name + authorseach.author.org] = number;
                    webBrowser1_DocumentCompleted(null, null);
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            connectdb();
            AuthortoDepartment();
        }
        private void AuthortoDepartment()
        {
            //检索所有关键词
            List<string> keywords = new List<string>();
            MySqlConnection m_Connection = this.connectdb();
            m_Connection.Open();
            string cp = "select keyword from searchword";
            MySqlCommand mysqlcom = new MySqlCommand(cp, m_Connection);
            MySqlDataReader mysqlread = mysqlcom.ExecuteReader(CommandBehavior.CloseConnection);
            while (mysqlread.Read())
            {
                string temp = mysqlread.GetString(mysqlread.GetOrdinal("keyword"));
                keywords.Add(temp);
            }
            mysqlread.Close();
            m_Connection.Close();
            foreach (string keyword in keywords)
            {
                AuthortoDepartmen_Eachkeyword(keyword);
            }
            int i = 0;
        }
        private void AuthortoDepartmen_Eachkeyword(string keyword)
        {
            searchneeds.Clear();
            MySqlConnection m_Connection = this.connectdb();
            m_Connection.Open();
            List<Paper> papers = new List<Paper>();
            string cp = "select orgn,paper_id,authors from paper where searchword like '" + keyword + '\'';
            MySqlCommand mysqlcom = new MySqlCommand(cp, m_Connection);
            MySqlDataReader mysqlread = mysqlcom.ExecuteReader(CommandBehavior.CloseConnection);
            while (mysqlread.Read())
            {
                Paper paper = new Paper();
                try
                {
                    paper.paperid = mysqlread.GetString(mysqlread.GetOrdinal("paper_id"));
                    paper.authors = mysqlread.GetString(mysqlread.GetOrdinal("authors"));
                    paper.orgs = mysqlread.GetString(mysqlread.GetOrdinal("orgn"));
                    papers.Add(paper);
                }
                catch (Exception e)
                {
                    ;
                }                
            }
           
            if (papers.Count < 1)
            {
                return;
            }

            //开始处理

            int authorid = 1;
            List<PaperAuthor> paperauthors = new List<PaperAuthor>();
            List<Author> authors = new List<Author>();
            HashSet<string> searstr = new HashSet<string>();
           

            foreach (Paper paper in papers)
            {
                string[] orgstrs = paper.orgs.Split(';');
                List<string> orglist = orgstrs.ToList();
                int length = orgstrs.Length;
                for (int i = 0; i < length; i++)
                {
                    string temp=checkString(orgstrs[i]) ;
                    if (temp== "")
                    {
                        orglist.Remove(orgstrs[i]);
                        //length = length - 1;
                    }
                    else
                    {
                        orglist[orglist.IndexOf(orgstrs[i])] = temp;
                    }
                
                }
                orgstrs = orglist.ToArray();


                string[] authorstrs = paper.authors.Split(';');
                if (orgstrs.Length == 1)
                {
                    //如果只有一个单位，直接放入作者表，并且建立作者字典
                    foreach (string authorstr in authorstrs)
                    {
                        PaperAuthor paperauthor = new PaperAuthor();
                        Author author = this.FindAuthorID(authorstr, orgstrs[0], authorid, authors);//判断并输出作者                        
                        if (author.authorid - authorid ==0)
                        {//如果是新作者，
                            authorid = authorid+1;//ID更新
                            authors.Add(author);//添加到author set  
                        }
                        paperauthor.authorid = author.authorid;
                        paperauthor.authorname = author.name;
                        paperauthor.org = author.org;
                        paperauthor.paperid = paper.paperid;
                        paperauthors.Add(paperauthor);
                    }
                }
            }
            //处理多个单位的情况
            foreach (Paper paper in papers)
            {
                string[] orgstrs = paper.orgs.Split(';');
                List<string> orglist = orgstrs.ToList();
                int length = orgstrs.Length;
                for (int i = 0; i < length; i++)
                {
                    string temp = checkString(orgstrs[i]);
                    if (temp == "")
                    {
                        orglist.Remove(orgstrs[i]);
                        //length = length - 1;
                    }
                    else
                    {
                        orglist[orglist.IndexOf(orgstrs[i])] = temp;
                    }

                }
                orgstrs = orglist.ToArray();
               
                string[] authorstrs = paper.authors.Split(';');  
                List<string> orgstrslist=orgstrs.ToList<string>();
                List<string> authorstrslist=authorstrs.ToList<string>();

                int orglength = orgstrs.Length;
                if (orglength > 1)
                {
                    //具体情况
                    foreach (string authorstr in authorstrs)
                    {
                        int flag = 0;
                        //查字典
                        List<Author> authorsamename = authors.Where(x =>string.Equals(x.name,authorstr)).ToList<Author>();
                        foreach (Author author in authorsamename)
                        {
                            string orgtemp = author.org;
                            foreach (string orgstr in orgstrs)
                            {
                                if (orgstr.IndexOf(orgtemp) + orgtemp.IndexOf(orgstr) > -2)//判断相同单位方式
                                {
                                    PaperAuthor paperauthor = new PaperAuthor();
                                    paperauthor.authorid = author.authorid;
                                    paperauthor.authorname = author.name;
                                    paperauthor.org = author.org;
                                    paperauthor.paperid = paper.paperid;
                                    paperauthors.Add(paperauthor);
                                    flag = 1;
                                    orglength = orglength - 1;
                                    orgstrslist.Remove(orgstr);
                                    continue;
                                }
                             }
                             if (flag==1)//字典查到了
                             {
                                    authorstrslist.Remove(authorstr);
                                    continue;
                             }
                           }
                    }
                    authorstrs=authorstrslist.ToArray();
                    foreach (string authorstr in authorstrs)
                    {                        
                        //如果只剩下一个作者和一个单位
                        if(authorstrslist.Count==1 && orgstrslist.Count==1)
                        {
                            Author authornew = new Author();
                            authornew.name = authorstr;
                            authornew.org = orgstrslist.ElementAt(0);
                            authornew.authorid = authorid++;
                            authors.Add(authornew);
                            PaperAuthor paperauthor = new PaperAuthor();
                            paperauthor.authorid = authornew.authorid;
                            paperauthor.authorname = authornew.name;
                            paperauthor.org = authornew.org;
                            paperauthor.paperid = paper.paperid;
                            paperauthors.Add(paperauthor);

                        }
                            //其他情况需要到CNKI上去检索
                        else
                       {
                           //
                            Author authornew = new Author();
                            authornew.name = authorstr;
                            authornew.org = paper.orgs;//现在还没有给定地址
                            authornew.authorid = authorid++;
                            authors.Add(authornew);
                            PaperAuthor paperauthor = new PaperAuthor();
                            paperauthor.authorid = authornew.authorid;
                            paperauthor.authorname = authornew.name;
                            paperauthor.org = authornew.org;
                            paperauthor.paperid = paper.paperid;
                            paperauthors.Add(paperauthor);
                            foreach (string orgstr in orgstrs)
                            {
                                Author searchneed = new Author();
                                searchneed.authorid=authorid;
                                searchneed.name=authorstr;
                                searchneed.org=orgstr;
                                if (checkString(orgstr)!="")
                                {
                                    searchneeds.Add(searchneed);
                                }
                            }
                        }
                    }

                  }
                }
            //进行检索           
            SearcherFromCNKI();
            //HashSet<Author> authorsnew = new HashSet<Author>();
            //authorsnew = authors;

            for (int i = 0; i < authors.Count;i++ )
            {
                Author author = new Author();
                author.authorid = authors[i].authorid;
                author.name = authors[i].name;
                author.org = ""; int num = 0;
                foreach (AuthorSearch authorsearch in searchneedsaved)
                {
                    if (authorsearch.author.authorid == authors[i].authorid)
                    {
                        int papernumber=authorsearch.number;
                        if (papernumber > num)
                        {
                            author.org = authorsearch.author.org;
                        }
                    }
                }
                authors[i] = author;
              }
            writetoDb(authors,paperauthors);
        }        
        private Author FindAuthorID(string authorstr, string orgstr, int authorid, List<Author> authors)
    {        
        foreach (Author author in authors)
        {
            if (string.Equals(author.name, authorstr))
            {
                if (author.org.IndexOf(orgstr) + orgstr.IndexOf(author.org)>-2)
                {                    
                    return author;
                }
            }
        }
        Author authornew = new Author();
        authornew.name = authorstr;
        authornew.org = orgstr;
        authornew.authorid = authorid;
        return authornew;
    }

        private void SearcherFromCNKI()
        {
            string cnki_search_url = "http://epub.cnki.net/kns/brief/result.aspx?dbprefix=scdb&action=scdbsearch&db_opt=SCDB";
            //searchnameold = searchnamenew;
            //searchnamenew = new HashSet<CAuthor>();
            webBrowser1.ScriptErrorsSuppressed = true;
            //     webBrowser1.NewWindow += new CancelEventHandler(webBrowser1_NewWindow);
            //     Web_V1 = (SHDocVw.WebBrowser_V1)this.webBrowser1.ActiveXInstance;
            //     Web_V1.NewWindow += new SHDocVw.DWebBrowserEvents_NewWindowEventHandler(Web_V1_NewWindow);
            webBrowser1.Navigate(cnki_search_url);
            while (loading)
            {
                Application.DoEvents();
            }
        }

        private void searchforName_Enp(string name, string enp)
        {
            //state = 1;
            System.Windows.Forms.HtmlDocument doc = webBrowser1.Document;
            int m_state = 0;
            foreach (HtmlElement em in doc.All) //轮循 
            {
                string str = em.Name;
                string id = em.Id;
                switch (id)
                {
                    case "au_1_value1":
                        m_state = 1;
                        em.SetAttribute("value", name);
                        break;
                    case "au_1_value2":
                        if (m_state != 1)
                            return;
                        em.SetAttribute("value", enp);
                        break;
                    case "btnSearch":
                        state = 2;
                        em.InvokeMember("onclick");
                        //getPapers();
                        break;
                }
            }
        }

        private int getPapers()
        {
            state = 0;
            int temp = 0;
            //System.Threading.Thread.Sleep(5000);
            System.Windows.Forms.HtmlDocument doc = webBrowser1.Document.Window.Frames["iframeResult"].Document;           
            foreach (HtmlElement em in doc.All) //轮循 
            {

                if (em.GetAttribute("classname") == "pagerTitleCell")
                {
                    string numbers = em.OuterText;
                    Match mathc = Regex.Match(numbers, "\\d+");
                    temp = int.Parse(mathc.Value);
                    return temp;
                }
            }
            return temp;
        }

        private string checkString(string str)
        { 
            //判断字符串是否有一个
            //bool flag = false;

            if(str == "")
            {                
                return "";
            }

            Match match;
            match = Regex.Match(str, @"^([\u4E00-\u9FFF]+)");

            if (match.Length < 0)
            {
                return "";
            }

            if (match.Value.Length < 5)
            {
                return "";
            }
            
            return match.Value;
        
        }

        private void writetoDb(List<Author> authors, List<PaperAuthor> paperauthors)
        { 
            
            
        }
    }
}
