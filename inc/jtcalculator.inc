<script type="text/javascript"><!--
/* Script by: www.jtricks.com
 * Version: 20060426
 * Latest version:
 * www.jtricks.com/javascript/forms/calculator.html
 */
var calc_image = '/icormanager/images/icon_calculator.png';

var calc_start_num = false;
var calc_cur_oper = '';
var calc_acc = 0;
var calc_prev_oper = 0;
var calc_prev_entry = 0;

var calc_field;

function calc_num_value()
{
    if (calc_field.value == '0.')
        return calc_field.value;
    else if (isNaN(calc_field.value))
        return 0;
    else
        return parseFloat(calc_field.value);
}

function calc_perform_oper(oper, val1, val2)
{
    switch(oper)
    {
    case '*':
        return val1 * val2;
    case '/':
        return val1 / val2;
    case '+':
        return Math.round(
                val1 * 10000000000000 +
                val2 * 10000000000000) /
               10000000000000;
    case '-':
        return Math.round(
                val1 * 10000000000000 -
                val2 * 10000000000000) /
               10000000000000;
    default:
        return val2;
    }
}

function calc_press(button)
{
    switch (button)
    {
    case 'CE':
        calc_start_num = true;
        calc_field.value = "0";
        return;

    case 'C':
        calc_acc = 0;
        calc_prev_entry = 0;
        calc_cur_oper = "";
        calc_prev_oper = "";
        calc_start_num = true;
        calc_field.value = "0";
        return;

    case '+/-':
        calc_field.value = calc_num_value() * -1;
        return;

    case ',':
    case '.':
        if (calc_start_num || isNaN(calc_field.value))
        {
            calc_field.value = "0.";
            calc_start_num = false;
        }
        else
        {
            if (calc_field.value.indexOf(".") == -1)
                calc_field.value += ".";
        }
        return;

    case '+':
    case '-':
    case '/':
    case '*':
        calc_prev_oper = button;

        if (calc_start_num)
        {
            calc_cur_oper = button;
            return;
        }

        // No break here.

    case '=':
        calc_start_num = true;

        if (button == '=' && calc_cur_oper != '=')
        {
            calc_prev_entry = calc_num_value();
        }

        if (button == '=' && calc_cur_oper == '=')
        {
            calc_acc = calc_perform_oper(
                calc_prev_oper, calc_acc, calc_prev_entry)
        }
        else
            calc_acc = calc_perform_oper(
                calc_cur_oper, calc_acc, calc_num_value());

        calc_field.value = calc_acc;
        calc_cur_oper = button;

        return;
    }

    if (calc_start_num)
    {
        calc_field.value  = button;
        calc_start_num = false;
    }
    else
    {
        if (calc_num_value() == "0")
            calc_field.value = button;
        else
            calc_field.value += button;
    }
}

function calc_setup(frm, inp)
{
    if (!document.getElementById)
        return;

    document.write(
        '<a href="javascript:void(0);" ' +
        'onclick="calc_show_hide(\'' +
        frm + '\', \'' + inp +
        '\')"><img  style="border:none" alt="kalkulator" ' +
        'src="' + calc_image +
        '"><' + '/a>');
}

function move_box(an, box)
{
    var cleft = 0;
    var ctop = 0;
    var obj = an;

    while (obj.offsetParent)
    {
        cleft += obj.offsetLeft;
        ctop += obj.offsetTop;
        obj = obj.offsetParent;
    }

    box.style.left = cleft + 'px';

    ctop += an.offsetHeight + 2;

    // Handle Internet Explorer body margins,
    // which affect normal document, but not
    // absolute-positioned stuff.
    if (document.body.currentStyle &&
        document.body.currentStyle['marginTop'])
    {
        ctop += parseInt(
            document.body.currentStyle['marginTop']);
    }

    box.style.top = ctop + 'px';
}

function calc_show_hide(frm, inp)
{
    var boxcalc = document.getElementById('attach_calc_div');
    var to_obj = document.forms[frm].elements[inp];

    if (boxcalc == null)
        return;

    if (to_obj != calc_field ||
        boxcalc.style.display=='none')
    {
        // Show and move calculator.

        move_box(to_obj, boxcalc);
        calc_field = to_obj;
        boxcalc.style.display='block';
    }
    else
        // Hide currently shown calculator.
        boxcalc.style.display='none';
    return false;
}

function calc_hide()
{
    document.getElementById('attach_calc_div')
        .style.display = 'none';
}

//--></script>

<div id="attach_calc_div"
 style="background:white;padding:2px;border:1px solid #2266AA;position:absolute;display:none">
<form name="attach_calc_form" action=""
 style="margin-bottom:0px;margin-top:0px">

<table style="border:none;padding:0px;border-spacing:0px">

<tbody>

<tr><td></td><td></td><td></td><td></td>
<td align="center"><a style="font-size:9px;color:#2266AA"
 href="javascript:void(0);" onmouseup="calc_hide()">Zamknij</a></td>
</tr>

<tr>
<td><input style="width:40px" name="cbtn7"
 type="button" value="7" onmouseup="calc_press(7)"></td>
<td><input style="width:40px" name="cbtn8"
 type="button" value="8" onmouseup="calc_press(8)"></td>
<td><input style="width:40px" name="cbtn9"
 type="button" value="9" onmouseup="calc_press(9)"></td>
<td><input style="width:40px" name="cbtnp"
 type="button" value="+" onmouseup="calc_press('+')"></td>
<td><input style="width:40px" name="cbtnm"
 type="button" value="-" onmouseup="calc_press('-')"></td>
</tr>

<tr>
<td><input style="width:40px" name="cbtn4"
 type="button" value="4" onmouseup="calc_press(4)"></td>
<td><input style="width:40px" name="cbtn5"
 type="button" value="5" onmouseup="calc_press(5)"></td>
<td><input style="width:40px" name="cbtn6"
 type="button" value="6" onmouseup="calc_press(6)"></td>
<td><input style="width:40px" name="cbtnm"
 type="button" value="*" onmouseup="calc_press('*')"></td>
<td><input style="width:40px" name="cbtnd"
 type="button" value="/" onmouseup="calc_press('/')"></td>
</tr>

<tr>
<td><input style="width:40px" name="cbtn1"
 type="button" value="1" onmouseup="calc_press(1)"></td>
<td><input style="width:40px" name="cbtn2"
 type="button" value="2" onmouseup="calc_press(2)"></td>
<td><input style="width:40px" name="cbtn3"
 type="button" value="3" onmouseup="calc_press(3)"></td>
<td><input style="width:40px" name="cbtnc"
 type="button" value="C" onmouseup="calc_press('C')"></td>
<td><input style="width:40px" name="cbtne"
 type="button" value="CE" onmouseup="calc_press('CE')"></td>
</tr>

<tr>
<td><input style="width:40px" name="cbtn0"
 type="button" value="0" onmouseup="calc_press(0)"></td>
<td><input style="width:40px" name="cbtnd"
 type="button" value="." onmouseup="calc_press('.')"></td>
<td><input style="width:40px" name="cbtnn"
 type="button" value="+/-" onmouseup="calc_press('+/-')"></td>
<td colspan=2><input style="width:84px" name="cbtnq"
 type="Button" value="=" onmouseup="calc_press('=')"></td>
</tr>

</table>
</form>
</div>
