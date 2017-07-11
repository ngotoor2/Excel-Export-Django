from django.forms import ModelForm, ValidationError
from django.http import HttpResponse, HttpResponseRedirect
from django.shortcuts import render_to_response
from django.template import RequestContext
import hashlib, xlwt, tempfile, os
from django.contrib.contenttypes.models import ContentType
from django.contrib.auth.decorators import login_required
from django.contrib import admin

@login_required
def admin_export_xls(request, app, model):
    mc = ContentType.objects.get(app_label=app, model=model).model_class()
    
    wb = xlwt.Workbook()
    ws = wb.add_sheet(unicode(mc._meta.verbose_name_plural))
    for i, f in enumerate(mc._meta.fields):
        ws.write(0,i, f.verbose_name)
    qs = mc.objects.all()
    for ri, row in enumerate(qs):
        for ci, f in enumerate(mc._meta.fields):
            # Let us test if the attribute is integer field and
            # if so do NOT use unicode
            if f.get_internal_type() == "IntegerField":
                ws.write(ri+1, ci, getattr(row, f.name))
            else:
                ws.write(ri+1, ci, unicode(getattr(row, f.name)))
    fd, fn = tempfile.mkstemp()
    os.close(fd)
    wb.save(fn)
    fh = open(fn, 'rb')
    resp = fh.read()
    fh.close()
    response = HttpResponse(resp, mimetype='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="%s.xls"' % (unicode(mc._meta.verbose_name_plural),)

    return response

class Utility():
    def __init__(self):
        self.sum = 0
        self.num = 0
    
def export_xls(modeladmin, request, queryset):
    meta = modeladmin.model._meta
    unicode_meta = unicode(meta)
        
    def get_verbose_name(fieldname):
        name = filter(lambda x: x.name == fieldname, meta.fields)
        if name:
            return (name[0].verbose_name or name[0].name).upper()
        return fieldname.upper()
    
    wbk = xlwt.Workbook()
    sht = wbk.add_sheet(unicode(meta.verbose_name_plural))

    if unicode_meta == 'cooking101.c101_signup': # For cooking101 app, weed out the introduced database fields
        
        li_e = list(enumerate(meta.fields))
        li = []
        li_meta = []
        for l in li_e:
            li_meta.append(l[1])
            li.append(l[1].verbose_name)
            
        li = li[:len(li)-3]
        li_meta = li_meta[:len(li_meta)-3]
        for l in li:
            sht.write(0, li.index(l), l) # first row

        li_e1 = list(enumerate(queryset))
        li1=[]
        for l in li_e1:
            li1.append(l[1])
        li2 = []
        for l in li1:
            for m in li_meta:
                sht.write(li1.index(l)+1, li_meta.index(m), unicode(getattr(l, m.name))) # subsequent rows


    else:
        for j, fieldname in enumerate(meta.fields):
            
            sht.write(0, j, fieldname.verbose_name)

        for i, row in enumerate(queryset):
            for j, fieldname in enumerate(meta.fields):
                if fieldname.get_internal_type() == "IntegerField":
                    sht.write(i+1, j, getattr(row, fieldname.name))
                else:
                    sht.write(i + 1, j, unicode(getattr(row, fieldname.name)))
                    
            
    fd, fn = tempfile.mkstemp()
    os.close(fd)
    wbk.save(fn)
    fh = open(fn, 'rb')
    resp = fh.read()
    fh.close()
    response = HttpResponse(resp, mimetype='application/ms-excel')
    response['Content-Disposition'] = 'attachment; filename="%s.xls"' % (unicode(meta.verbose_name_plural),)
    
    return response
export_xls.short_description = "Export filtered items to XLS (select all and export)"
admin.site.add_action(export_xls)
