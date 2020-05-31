package com.example.fragmentest;

import androidx.appcompat.app.AppCompatActivity;
import androidx.fragment.app.ListFragment;

import android.os.Bundle;
import android.widget.Toast;

import java.util.ArrayList;

public class MainActivity extends AppCompatActivity {
    ArrayList<String> titlelist=new ArrayList<>();
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
       for(int i=0;i<20;i++){
           titlelist.add(String.valueOf(i));
       }
        setContentView(R.layout.activity_main);
        Listfragment listfragment=new Listfragment(titlelist,this);
        getSupportFragmentManager().beginTransaction().add(R.id.container,listfragment).commit();
        listfragment.setOnitemclickListener(new Listfragment.OnitemClickListener() {
            @Override
            public void onitemclick(int position) {
                Toast.makeText(MainActivity.this,"you click"+String.valueOf(position),Toast.LENGTH_LONG).show();
            }
        });
    }
}
