package com.gs2e.xlxreadapi.Service;

import com.gs2e.xlxreadapi.Model.Electeur;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.util.List;

public interface ElecteurExcelService {

    public void save(MultipartFile file) ;

    public ByteArrayInputStream load();

    public List<Electeur> getAllElecteurs() ;
}