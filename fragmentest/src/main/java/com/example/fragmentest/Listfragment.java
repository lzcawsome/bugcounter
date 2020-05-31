package com.example.fragmentest;

import android.content.Context;
import android.os.Bundle;

import androidx.annotation.NonNull;
import androidx.fragment.app.Fragment;
import androidx.fragment.app.ListFragment;

import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.BaseAdapter;
import android.widget.ImageView;
import android.widget.ListView;
import android.widget.TextView;

import java.util.ArrayList;


/**
 * A simple {@link Fragment} subclass.
 * Use the {@link Listfragment#newInstance} factory method to
 * create an instance of this fragment.
 */
public class Listfragment extends ListFragment {
    ArrayList<String> titlelist=new ArrayList<>();
    Context context;

    public Listfragment(ArrayList<String> titlelist, Context context) {
        this.titlelist = titlelist;
        this.context = context;
    }

    OnitemClickListener onitemClickListener;
    // TODO: Rename parameter arguments, choose names that match
    // the fragment initialization parameters, e.g. ARG_ITEM_NUMBER
    private static final String ARG_PARAM1 = "param1";
    private static final String ARG_PARAM2 = "param2";

    // TODO: Rename and change types of parameters
    private String mParam1;
    private String mParam2;

    public Listfragment() {
        // Required empty public constructor
    }

    /**
     * Use this factory method to create a new instance of
     * this fragment using the provided parameters.
     *
     * @param param1 Parameter 1.
     * @param param2 Parameter 2.
     * @return A new instance of fragment listfragment.
     */
    // TODO: Rename and change types and number of parameters
    public static Listfragment newInstance(String param1, String param2) {
        Listfragment fragment = new Listfragment();
        Bundle args = new Bundle();
        args.putString(ARG_PARAM1, param1);
        args.putString(ARG_PARAM2, param2);
        fragment.setArguments(args);
        return fragment;
    }

    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        if (getArguments() != null) {
            mParam1 = getArguments().getString(ARG_PARAM1);
            mParam2 = getArguments().getString(ARG_PARAM2);
        }
        setListAdapter(new Myadapter());
}

    @Override
    public void onListItemClick(@NonNull ListView l, @NonNull View v, int position, long id) {
        onitemClickListener.onitemclick(position);
    }

    public void setOnitemclickListener(OnitemClickListener onitemclickListener){
        this.onitemClickListener=onitemclickListener;
    }


 interface OnitemClickListener{
        public void onitemclick(int position);
 }



    class Myadapter extends BaseAdapter{

        @Override
        public int getCount() {
            return titlelist.size();
        }

        @Override
        public Object getItem(int position) {
            return titlelist.get(position);
        }

        @Override
        public long getItemId(int position) {
            return position;
        }

        @Override
        public View getView(int position, View convertView, ViewGroup parent) {
            View view = LayoutInflater.from(context).inflate(R.layout.item, null);
            TextView tv=view.findViewById(R.id.tv);
            tv.setText(titlelist.get(position));
            ImageView imageView=view.findViewById(R.id.img);
            imageView.setImageResource(R.mipmap.ic_launcher);
            return view;
        }
    }
}
