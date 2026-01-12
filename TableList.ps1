param
(
$FileName
)
   
$ReturnArray=@()

#Add new tables below:

if($filename -like "*CRT_UK_GEN_TEL_SKILL*")
{
    $ReturnArray = "sp_t_uk_gen_tel_skill_update"
}

elseif($filename-like "*CRT_UK_GEN_TEL_AGENT*") 
{
    $ReturnArray = "sp_t_uk_gen_tel_agent_update"
}

elseif($filename -like "*CRT_UK_GEN_TEL_OB*") 
{           
    $ReturnArray = "sp_t_uk_gen_tel_ob_update"
}

elseif($filename -like "*CRT_UK_GEN_TEL_LOGIN_LOGOUT*") 
{          
    $ReturnArray = "sp_t_uk_gen_tel_login_logout_update"
}

elseif($filename -like "*CRT_UK_GEN_TEL_LOGIN_MAIN_SKILL*") 
{            
    $ReturnArray = "sp_t_uk_gen_tel_login_main_skill_update"
}

elseif($filename -like "*CRT_UK_GEN_TEL_VDN*") 
{            
    $ReturnArray = "sp_t_uk_gen_tel_vdn_update"
}

elseif($filename -like "CRT_UK_GEN_FTE_OLD*") 
{
    $ReturnArray = "sp_t_uk_gen_fte_update"
}

elseif($filename -like "*CRT_UK_GEN_KRONOS*") 
{     
    $ReturnArray = "sp_t_uk_gen_kronos_update"
}

elseif($filename -like "*CRT_UK_GEN_MANAGER_HISTORY*") 
{        
    $ReturnArray = "sp_t_uk_gen_manager_history_update"
}

elseif($filename -like "*CRT_UK_S_MAP_UK_SKILL*") 
{          
    $ReturnArray = "sp_s_map_uk_skill_update"
}

elseif($filename -like "*Agent_Performance*") 
{          
    $ReturnArray = "sp_t_uk_gen_siox_agent_update"
}

elseif($filename -like "*Agent_Queues*") 
{         
    $ReturnArray = "sp_t_uk_gen_siox_queue_update"
}

elseif($filename -like "*Survey VoC*") 
{          
    $ReturnArray = "sp_t_uk_cs_jl_voc_update"
}

elseif($filename -like "*VBM KPI*") 
{            
    $ReturnArray = "sp_t_uk_cs_jl_VBM_update"
}

elseif($filename -like "*Callbacks*") 
{           
    $ReturnArray = "sp_t_uk_cs_jl_callbacks_update"
}

elseif($filename -like "*fcr*") 
{         
    $ReturnArray = "sp_t_uk_cs_jl_fcr_update"
}

elseif($filename -like "*forecasts*") 
{   
    $ReturnArray = "sp_t_uk_gen_forecasts_iex_update"
}

elseif($filename -like "*mailbox*") 
{      
    $ReturnArray = "sp_t_uk_cs_jl_mbox_update"
}

elseif($filename -like "*Complaint Status*") 
{      
    $ReturnArray = "sp_t_uk_cs_jl_cmpl_update"
}

elseif($filename -like "*John Lewis SMS*") 
{          
    $ReturnArray = "sp_t_uk_cs_jl_ivr_sms_update"
}

elseif($filename -like "*heathrow backoffice*") 
{          
    $ReturnArray = "sp_t_uk_cs_heathrow_bo_update"
}  

elseif($filename -like "*CRT_UK_GEN_TEL_INTERVAL*") 
{          
    $ReturnArray = "sp_t_uk_gen_tel_interval_update"
}  
elseif($filename -like "*CRT_UK_GEN_FTE_EMPLOYEES_MAIN*") 
{          
    $ReturnArray = "sp_t_uk_gen_fte_main_update"
}  
elseif($filename -like "*CRT_UK_GEN_FTE_EMPLOYEES_CLIENT_LOGIN*") 
{          
    $ReturnArray = "sp_t_uk_gen_fte_client_login_update"
}
elseif($filename -like "*CRT_UK_GEN_FTE_EMPLOYEES_CRM_LOGIN*") 
{          
    $ReturnArray = "sp_t_uk_gen_fte_crm_login_update"
}  
elseif($filename -like "*CRT_UK_GEN_FTE_EMPLOYEES_PHONE_LOGIN*") 
{          
    $ReturnArray = "sp_t_uk_gen_fte_phone_login_update"
}   

elseif($filename -like "*CRT_UK_GEN_FTE_EMPLOYEES_WIN_LOGIN*") 
{          
    $ReturnArray = "sp_t_uk_gen_fte_windows_login_update"
}   
elseif($filename -like "*CRT_UK_GEN_FTE_JL_RM_LOGIN*") 
{          
    $ReturnArray = "sp_t_uk_gen_fte_jl_rm_login_update"
} 
elseif($filename -like "*CRT_UK_GEN_FTE_HP_LOGIN*") 
{          
    $ReturnArray = "sp_t_uk_gen_fte_hp_login_update"
}     
elseif($filename -like "*JnJ Log Data*") 
{          
    $ReturnArray = "sp_t_uk_cs_jnj_bo"
} 
elseif($filename -like "*incident and complaint*") 
{          
    $ReturnArray = "sp_t_uk_cs_jl_incidents_weekly_update"
} 
elseif($filename -like "*SLA Performance*") 
{          
    $ReturnArray = "sp_t_uk_cs_flybe_sla_update"
}
elseif($filename -like "*Transaction Volume*") 
{          
    $ReturnArray = "sp_t_uk_cs_flybe_trans_vol_update"
}
elseif($filename -like "*CRT_UK_GEN_FTE_EMPLOYEES_CITRIX_LOGIN*") 
{          
    $ReturnArray = "sp_t_uk_gen_fte_citrix_login_update"
}    
return $ReturnArray