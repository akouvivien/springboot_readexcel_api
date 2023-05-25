package com.gs2e.xlxreadapi.ServiceImpl;

import com.gs2e.xlxreadapi.Model.Electeur;
import com.gs2e.xlxreadapi.Helper.ExcelHelper;
import com.gs2e.xlxreadapi.Repository.ElecteurRepository;
import com.gs2e.xlxreadapi.Service.ElecteurExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;
@Service
public class ElecteurExcelServiceImpl implements ElecteurExcelService {

    @Autowired
    ElecteurRepository electeurRepo;

    @Override
    public void save(MultipartFile file) {
        try {
            List<Electeur> electeurs = ExcelHelper.excelToElecteurs(file.getInputStream());
            electeurRepo.saveAll(electeurs);
        } catch (IOException e) {
            throw new RuntimeException("fail to store excel data: " + e.getMessage());
        }
}
    @Override
    public ByteArrayInputStream load() {
        List<Electeur> electeurs = electeurRepo.findAll();

        ByteArrayInputStream in = ExcelHelper.electeursToExcel(electeurs);
        return in;
    }
    @Override
    public List<Electeur> getAllElecteurs() {
        return electeurRepo.findAll();
    }
}
