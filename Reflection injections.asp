public class ExampleController : Controller
{
    public IActionResult Apply(string EffectName)
    {
        var EffectInstance  = Activator.CreateInstance(null, EffectName); // Noncompliant
        object EffectPlugin = EffectInstance.Unwrap();

        if ( ((IEffect)EffectPlugin).ApplyFilter() )
        {
            return Ok();
        }
        else
        {
            return Problem();
        }
    }
}

public interface IEffect
{
    bool ApplyFilter();
}