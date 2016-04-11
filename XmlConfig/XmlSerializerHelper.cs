using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Xml.Serialization;
using System.IO;




namespace ElancoPimsDdsParser
{
    

    public static class XmlSerializerHelper
    {
        public const string XmlConfigDirectory = "XmlConfig";

        public static void SerializeTuples(this List<Tuple<string, string>> listTuples, string fileName)
        {
            List<SerializableTuple<string, string>> listSerialTpls = listTuples.ConvertAll(
                new Converter<Tuple<string, string>, SerializableTuple<string, string>>(TupleToSerializableTuple));

            XmlSerializer serializer = new XmlSerializer(typeof(List<SerializableTuple<string, string>>));

            TextWriter writer = new StreamWriter(fileName);
            try
            {
                serializer.Serialize(writer, listSerialTpls);
            }
            catch (Exception e)
            {
                Console.WriteLine(String.Format("ERROR - Unable to deserialize {0}:{1}", fileName, e.ToString()));
            }

            writer.Close();
        }

        public static List<Tuple<string, string>> DeserializeTuples(string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<SerializableTuple<string, string>>));
            FileStream fs = new FileStream(fileName, FileMode.Open);
            List<SerializableTuple<string, string>> retreivedList = (List<SerializableTuple<string, string>>)(serializer.Deserialize(fs));

            List<Tuple<string, string>> listTpls = retreivedList.ConvertAll(
                new Converter<SerializableTuple<string, string>, Tuple<string, string>>(SerializableTupleToTuple));

            return listTpls;
        }


        public class SerializableTuple<T1, T2>
        {
            public T1 Item1 { get; set; }
            public T2 Item2 { get; set; }

            public static implicit operator Tuple<T1, T2>(SerializableTuple<T1, T2> st)
            {
                return Tuple.Create(st.Item1, st.Item2);
            }

            public static implicit operator SerializableTuple<T1, T2>(Tuple<T1, T2> t)
            {
                return new SerializableTuple<T1, T2>()
                {
                    Item1 = t.Item1,
                    Item2 = t.Item2
                };
            }

            public SerializableTuple() { }
        }

        /* delegate Converter */
        /// <summary>
        /// Delegate for converting Tuple<string,string> to SerializableTuple<string,string>
        /// </summary>
        /// <param name="tpl"></param>
        /// <returns></returns>
        public static SerializableTuple<string, string> TupleToSerializableTuple(Tuple<string, string> tpl)
        {
            SerializableTuple<string, string> serializableTuple = tpl;
            return serializableTuple;
        }

        /// <summary>
        /// Delegate for converting SerializableTuple<string,string> to Tuple<string,string>
        /// </summary>
        /// <param name="serTpl"></param>
        /// <returns></returns>
        public static Tuple<string, string> SerializableTupleToTuple(SerializableTuple<string, string> serTpl)
        {
            Tuple<string, string> tpl = serTpl;
            return tpl;
        }

    } /* end of XMLHelper static class */

    /********************/

    public static class XmlSerializerHelperV2
    {
        public static void SerializeTuples(this List<Tuple<string, string, int>> listTuples, string fileName)
        {
            List<SerializableTuple<string, string, int>> listSerialTpls = listTuples.ConvertAll(
                new Converter<Tuple<string, string, int>, SerializableTuple<string, string, int>>(TupleToSerializableTuple));

            XmlSerializer serializer = new XmlSerializer(typeof(List<SerializableTuple<string, string, int>>));

            TextWriter writer = new StreamWriter(fileName);
            try
            {
                serializer.Serialize(writer, listSerialTpls);
            }
            catch (Exception e)
            {
                Console.WriteLine(String.Format("ERROR - Unable to deserialize {0}:{1}", fileName, e.ToString()));
            }

            writer.Close();
        }

        public static List<Tuple<string, string, int>> DeserializeTuples(string fileName)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(List<SerializableTuple<string, string, int>>));
            FileStream fs = new FileStream(fileName, FileMode.Open);
            List<SerializableTuple<string, string, int>> retreivedList = (List<SerializableTuple<string, string, int>>)(serializer.Deserialize(fs));

            List<Tuple<string, string, int>> listTpls = retreivedList.ConvertAll(
                new Converter<SerializableTuple<string, string, int>, Tuple<string, string, int>>(SerializableTupleToTuple));

            return listTpls;
        }


        public class SerializableTuple<T1, T2, T3>
        {
            public T1 Item1 { get; set; }
            public T2 Item2 { get; set; }
            public T3 Item3 { get; set; }

            public static implicit operator Tuple<T1, T2, T3>(SerializableTuple<T1, T2, T3> st)
            {
                return Tuple.Create(st.Item1, st.Item2, st.Item3);
            }

            public static implicit operator SerializableTuple<T1, T2, T3>(Tuple<T1, T2, T3> t)
            {
                return new SerializableTuple<T1, T2, T3>()
                {
                    Item1 = t.Item1,
                    Item2 = t.Item2,
                    Item3 = t.Item3
                };
            }

            public SerializableTuple() { }
        }

        /* delegate Converter */
        /// <summary>
        /// Delegate for converting Tuple<string,string> to SerializableTuple<string,string>
        /// </summary>
        /// <param name="tpl"></param>
        /// <returns></returns>
        public static SerializableTuple<string, string, int> TupleToSerializableTuple(Tuple<string, string, int> tpl)
        {
            SerializableTuple<string, string, int> serializableTuple = tpl;
            return serializableTuple;
        }

        /// <summary>
        /// Delegate for converting SerializableTuple<string,string> to Tuple<string,string>
        /// </summary>
        /// <param name="serTpl"></param>
        /// <returns></returns>
        public static Tuple<string, string, int> SerializableTupleToTuple(SerializableTuple<string, string, int> serTpl)
        {
            Tuple<string, string, int> tpl = serTpl;
            return tpl;
        }

    } /* end of XMLHelper-V2 static class */

}
