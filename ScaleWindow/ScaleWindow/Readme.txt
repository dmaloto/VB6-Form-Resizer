Usage
1. Add the UserControl
Place ScalableForm directly on a form (not inside another container).

2. Set baseline dimensions (Optional)
In your form’s Form_Load, set the baseline dimensions to match the IDE design size:
Me.ScalableForm1.BaseFormWidth = 17280
Me.ScalableForm1.BaseFormHeight = 11475
This ensures scaling is calculated relative to the original design size.

3. Set scale percentage
By default, ScalableForm scales to 80% of the design size.
You can override this by calling: 
Me.ScalableForm1.SetScalePercentage = 90   ' Example: scale to 90%

4. Automatic behavior
On Form_Load, ScalableForm initializes and scales controls.
On Form_Resize, ScalableForm automatically adjusts all child controls proportionally.
Edge and corner grips allow manual resizing with proper cursors.

With this setup, your forms will always scale proportionally to the percentage you specify, while respecting the IDE design baseline.
