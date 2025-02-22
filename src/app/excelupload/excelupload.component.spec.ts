import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExceluploadComponent } from './excelupload.component';

describe('ExceluploadComponent', () => {
  let component: ExceluploadComponent;
  let fixture: ComponentFixture<ExceluploadComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ExceluploadComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(ExceluploadComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
