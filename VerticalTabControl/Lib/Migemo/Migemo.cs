using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;
using System.Runtime;

namespace VerticalTabControlLib.Migemo
{
    // https://github.com/koron/cmigemo/blob/master/tools/Migemo.cs
    public class Migemo : IDisposable
    {
        #region Enumerations (from migemo.h)
        #region enum DictionaryIndex
        public enum DictionaryId
        {
            Invalid = 0,
            Migemo = 1,
            RomaToHira = 2,
            HiraToKata = 3,
            HanToZen = 4,
        }
        #endregion

        #region enum OperatorIndex
        public enum OperatorIndex
        {
            Or = 0,
            NestIn = 1,
            NestOut = 2,
            SelectIn = 3,
            SelectOut = 4,
            NewLine = 5,
        }
        #endregion
        #endregion

        #region AnyCPU対応の為のプリロード 
        private const string DllName = "migemo.dll";

        [DllImport("kernel32")]
        private static extern IntPtr LoadLibrary(string lpFileName);

        static Migemo()
        {
            var folder = Path.Combine("Lib", "Migemo", Environment.Is64BitProcess ? "x64" : "x86");
            // ぬるぽで死ぬ場合がある？
            var dllPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), folder, DllName);

            var success = File.Exists(dllPath) && LoadLibrary(dllPath) != IntPtr.Zero;
            if (!success)
                throw new DllNotFoundException($"{dllPath}:dllが見つかりません");

        }

        #endregion

        #region Link to migemo.dll
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern IntPtr migemo_open(string dict);
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        private static extern void migemo_close(IntPtr obj);
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern IntPtr migemo_query(IntPtr obj, string query);
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        private static extern void migemo_release(IntPtr obj, IntPtr result);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern int migemo_set_operator(IntPtr obj,
                OperatorIndex index, string op);
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr migemo_get_operator(IntPtr obj,
                OperatorIndex index);

        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl, CharSet = CharSet.Ansi)]
        private static extern DictionaryId migemo_load(IntPtr obj,
                DictionaryId id, string file);
        [DllImport(DllName, CallingConvention = CallingConvention.Cdecl)]
        private static extern int migemo_is_enable(IntPtr obj);
        #endregion

        private IntPtr migemoObject = IntPtr.Zero;

        public IntPtr MigemoObject
        {
            get
            {
                return this.migemoObject;
            }
        }

        public bool SetOperator(OperatorIndex index, string op)
        {
            return migemo_set_operator(this.migemoObject, index, op) != 0;
        }
        public string GetOperator(OperatorIndex index)
        {
            IntPtr result = migemo_get_operator(this.migemoObject, index);
            if (result != IntPtr.Zero)
                return Marshal.PtrToStringAnsi(result);
            else
                return "";
        }

        #region Operator properties
        public string OperatorOr
        {
            get { return GetOperator(OperatorIndex.Or); }
            set { SetOperator(OperatorIndex.Or, value); }
        }
        public string OperatorNestIn
        {
            get { return GetOperator(OperatorIndex.NestIn); }
            set { SetOperator(OperatorIndex.NestIn, value); }
        }
        public string OperatorNestOut
        {
            get { return GetOperator(OperatorIndex.NestOut); }
            set { SetOperator(OperatorIndex.NestOut, value); }
        }
        public string OperatorSelectIn
        {
            get { return GetOperator(OperatorIndex.SelectIn); }
            set { SetOperator(OperatorIndex.SelectIn, value); }
        }
        public string OperatorSelectOut
        {
            get { return GetOperator(OperatorIndex.SelectOut); }
            set { SetOperator(OperatorIndex.SelectOut, value); }
        }
        public string OperatorNewLine
        {
            get { return GetOperator(OperatorIndex.NewLine); }
            set { SetOperator(OperatorIndex.NewLine, value); }
        }
        #endregion

        public bool LoadDictionary(DictionaryId id, string file)
        {
            DictionaryId result = migemo_load(this.migemoObject, id, file);
            return result == id;
        }

        public bool IsEnable()
        {
            return migemo_is_enable(this.migemoObject) != 0;
        }

        public Regex GetRegex(string query)
        {
            return new Regex(Query(query));
        }

        public string Query(string query)
        {
            IntPtr result = migemo_query(this.migemoObject, query);
            if (result != IntPtr.Zero)
            {
                string retval = Marshal.PtrToStringAnsi(result);
                migemo_release(this.migemoObject, result);
                return retval;
            }
            else
                return "";
        }

        public void Dispose()
        {
            //Console.WriteLine("HERE ("+this.migemoObject+")");
            if (this.migemoObject != IntPtr.Zero)
            {
                //Console.WriteLine("migemo_close() is called");
                migemo_close(this.migemoObject);
                this.migemoObject = IntPtr.Zero;
            }
            GC.SuppressFinalize(this);
        }

        public Migemo(string dictpath)
        {
            this.migemoObject = migemo_open(dictpath);
            this.OperatorNestIn = "(?:";
            //this.OperatorNewLine = "\\s*";
        }

        public Migemo() : this(null)
        {
        }

        ~Migemo()
        {
            this.Dispose();
        }

        private const string DefaultDictPath = "Lib/Migemo/dict/migemo-dict";
        public static Migemo GetDefault()
        {
            return new Migemo(DefaultDictPath);
        }

    }
}
