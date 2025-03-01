from django import forms
from .models import BaseDados

class ReservasForm(forms.ModelForm):
    
    class Meta:
        model = BaseDados
        # Substitua '...' com os nomes dos campos reais que você quer incluir no seu formulário
        fields = ['id', 'entrada', 'saida', 'nome', 'diarias', 'graduacao', 'telefone', 'qtde_quartos', 'qtde_hosp', 'especial', 'status_reserva', 'email', 'cpf', 'status', 'tipo', 'sexo', 'cidade', 'uf', 'nome_acomp1', 'vinculo_acomp1', 'idade_acomp1', 'sexo_acomp1', 'nome_acomp2', 'vinculo_acomp2', 'idade_acomp2', 'sexo_acomp2', 'nome_acomp3', 'vinculo_acomp3', 'idade_acomp3', 'sexo_acomp3', 'nome_acomp4', 'vinculo_acomp4', 'idade_acomp4', 'sexo_acomp4', 'nome_acomp5', 'vinculo_acomp5', 'idade_acomp5', 'sexo_acomp5', 'mhex', 'uh', 'motivo_viagem', 'valor_ajuste', 'observacao']
        widgets = {
            'entrada': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),
            'saida': forms.DateInput(attrs={'type': 'date', 'class': 'form-control'}),           
            
        }

    def __init__(self, *args, **kwargs):
        super(ReservasForm, self).__init__(*args, **kwargs)

        # Garante que as datas sejam carregadas corretamente no formato YYYY-MM-DD
        if self.instance and self.instance.pk:
            if self.instance.entrada:
                self.fields['entrada'].initial = self.instance.entrada.strftime('%Y-%m-%d')
            if self.instance.saida:
                self.fields['saida'].initial = self.instance.saida.strftime('%Y-%m-%d')      

        for field_name, field in self.fields.items():
            field.widget.attrs['style'] = 'height: 37px; border: 1px solid #d1d3e2;'
           
                      
            if field_name == 'qtde_hosp':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs.update({'id': 'id_qtde_hosp'})
                field.widget.attrs['required'] = 'required'
            if field_name == 'qtde_quartos':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputQtde_quartos'     

            if field_name == 'entrada':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'entrada'
                field.widget.attrs['required'] = 'required'
            if field_name == 'saida':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'saida'
                field.widget.attrs['required'] = 'required'

            if field_name == 'nome':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'nome'
                field.widget.attrs['required'] = 'required'
            if field_name == 'email':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'email' 
                field.widget.attrs['required'] = 'required'   
            if field_name == 'cpf':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'cpf'  
                field.widget.attrs['required'] = 'required'
            if field_name == 'telefone':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'telefone' 
                field.widget.attrs['required'] = 'required'      
            if field_name == 'sexo':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputSexo'
                field.widget.attrs['required'] = 'required'
            
            if field_name == 'status':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputStatus'
                field.widget.attrs['required'] = 'required'
            
            if field_name == 'graduacao':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputGraduacao' 
                field.widget.attrs['required'] = 'required'    
            if field_name == 'tipo':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputTipo' 
                field.widget.attrs['required'] = 'required'  
            
            if field_name == 'nome_acomp1':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputNome_acomp1' 
            if field_name == 'sexo_acomp1':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputSexo_acomp1' 
            if field_name == 'vinculo_acomp1':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputVinculo_acomp1'
            if field_name == 'idade_acomp1':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputIdade_acomp1'

            if field_name == 'nome_acomp2':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputNome_acomp2'    
            if field_name == 'sexo_acomp2':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputSexo_acomp2' 
            if field_name == 'vinculo_acomp2':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputVinculo_acomp2'
            if field_name == 'idade_acomp2':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputIdade_acomp2'
            
            if field_name == 'nome_acomp3':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputNome_acomp3'         
            if field_name == 'sexo_acomp3':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputSexo_acomp3' 
            if field_name == 'vinculo_acomp3':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputVinculo_acomp3'
            if field_name == 'idade_acomp3':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputIdade_acomp3'

            if field_name == 'nome_acomp4':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputNome_acomp4'
            if field_name == 'sexo_acomp4':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputSexo_acomp4' 
            if field_name == 'vinculo_acomp4':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputVinculo_acomp4'
            if field_name == 'idade_acomp4':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputIdade_acomp4'  

            if field_name == 'nome_acomp5':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputNome_acomp5'
            if field_name == 'sexo_acomp5':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputSexo_acomp5' 
            if field_name == 'vinculo_acomp5':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputVinculo_acomp5'
            if field_name == 'idade_acomp5':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputIdade_acomp5'    

            if field_name == 'especial':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputEspecial'
                field.widget.attrs['required'] = 'required'
            
            if field_name == 'cidade':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputCidade'  
                field.widget.attrs['required'] = 'required'    
            
            if field_name == 'uf':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputUF'
                field.widget.attrs['required'] = 'required'
                   
            if field_name == 'status_reserva':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'id_status_reserva'  


            if field_name == 'mhex':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputMhex'
            if field_name == 'uh':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'inputUH'

            if field_name == 'qtde_agua':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'qtde_agua'
            if field_name == 'qtde_refri':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'qtde_refri'
            if field_name == 'qtde_cerveja':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'qtde_cerveja' 
            if field_name == 'qtde_pet':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'qtde_pet' 


            if field_name == 'total_agua':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'total_agua' 
            if field_name == 'total_refri':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'total_refri' 
            if field_name == 'total_cerveja':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'total_cerveja'
            if field_name == 'total_pet':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'total_pet'    

            if field_name == 'total_consumacao':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'total_consumacao'   

            if field_name == 'forma_pagamento':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'forma_pagamento'  

            if field_name == 'nome_pagante':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'nome_pagante'                
            if field_name == 'cpf_pagante':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'cpf_pagante'  

            if field_name == 'motivo_viagem':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'motivo_viagem'  

            if field_name == 'desc_saude':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'desc_saude'   

            if field_name == 'observacao':
                field.widget.attrs['class'] = 'form-control'
                field.widget.attrs['id'] = 'observacao'                                  

