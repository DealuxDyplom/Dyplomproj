sealed class CustomBinder : SerializationBinder
{
   public override Type BindToType(string assemblyName, string typeName)
   {
       if (!(typeName == "type1" || typeName == "type2" || typeName == "type3"))
       {
          throw new SerializationException("Only type1, type2 and type3 are allowed");
       }
       return Assembly.Load(assemblyName).GetType(typeName);
   }
}

var myBinaryFormatter = new BinaryFormatter();
myBinaryFormatter.Binder = new CustomBinder();
myBinaryFormatter.Deserialize(stream);


public class CustomSafeTypeResolver : JavaScriptTypeResolver
{
   public override Type ResolveType(string id)
   {
      if(id != "ExpectedType") {
         throw new ArgumentNullException("Only ExpectedType is allowed during deserialization");
      }
      return Type.GetType(id);
   }
}

JavaScriptSerializer serializer = new JavaScriptSerializer(new CustomSafeTypeResolver());
serializer.Deserialize<ExpectedType>(json);